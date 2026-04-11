Attribute VB_Name = "acc_mod_Common_Utils"
'-------------------------------------
' Module: acc_mod_Common_Utils
' 説明： Access共通ユーティリティ（自動閉鎖通知・文字列正規化・高速化制御）
'-------------------------------------
Option Explicit

' --- Windows API 宣言 (Declarationsセクション：必ず一番上に配置) ---
' ※32bit版Accessでは Win64 の中身が赤字になりますが、#If で分岐しているため無視してOKです。
#If Win64 Then
    ' 64bit Access用
    Private Declare PtrSafe Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
        ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, _
        ByVal uType As Long, ByVal wLanguageID As Integer, ByVal dwMilliseconds As Long) As Long
#Else
    ' 32bit Access用
    Private Declare Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
        ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, _
        ByVal uType As Long, ByVal wLanguageID As Integer, ByVal dwMilliseconds As Long) As Long
#End If

' 一括実行中フラグ： Trueの場合、通知ポップアップを抑制する
Public IsBatchActive As Boolean

'--------------------------------------------
' プロシージャ名： Notify_AutoClose
' 概要： 指定した秒数で確実に閉じるポップアップ（Windows API版）
'--------------------------------------------
Public Sub Notify_AutoClose(ByVal msg As String, ByVal title As String, Optional ByVal sec As Long = 2)
    ' Python等のバッチ処理中はユーザー入力を妨げないよう表示しない
    If IsBatchActive Then Exit Sub

    ' 第4引数 64:情報アイコン / 第6引数:ミリ秒単位(sec * 1000)
    MessageBoxTimeout 0, msg, title, 64, 0, sec * 1000
End Sub
' ← プロシージャの終わり

'--------------------------------------------
' プロシージャ名： Notify_Smart_Popup
' 概要： 従来のWScript版（API版への移行推奨だが互換性のために維持）
'--------------------------------------------
Public Sub Notify_Smart_Popup(ByVal msg As String, ByVal title As String, Optional ByVal iconType As Integer = 64)
    If IsBatchActive Then Exit Sub
    
    On Error Resume Next
    CreateObject("WScript.Shell").PopUp msg, 1, title, iconType
    On Error GoTo 0
End Sub
' ← プロシージャの終わり

'--------------------------------------------
' 関数名： Normalize_Text
' 概要： 文字列の空白除去と半角統一を行う
'--------------------------------------------
Public Function Normalize_Text(ByVal txt As String) As String
    Dim result As String
    ' 全角・半角スペースを除去
    result = Replace(Replace(txt, " ", ""), "　", "")
    ' 英数字・記号を半角に統一
    On Error Resume Next
    result = StrConv(result, vbNarrow)
    On Error GoTo 0
    Normalize_Text = Trim(result)
End Function
' ← プロシージャの終わり

'--------------------------------------------
' 関数名： Get_ColumnIndex_Robust
' 概要： テーブルのタイトル名から列番号を取得する
'--------------------------------------------
Public Function Get_ColumnIndex_Robust(ByVal lo As Object, ByVal colName As String) As Long
    Dim col As Object
    Dim target As String: target = Normalize_Text(colName)
    
    For Each col In lo.ListColumns
        If Normalize_Text(col.Name) = target Then
            Get_ColumnIndex_Robust = col.Index
            Exit Function
        End If
    Next col
    Get_ColumnIndex_Robust = 0
End Function
' ← プロシージャの終わり

'--------------------------------------------
' プロシージャ名： Fast_Mode_Toggle
' 概要： 高速実行モードのON/OFF
'--------------------------------------------
Public Sub Fast_Mode_Toggle(ByVal isOn As Boolean, Optional ByVal targetExcelApp As Object = Nothing)
    ' Access側の制御
    DoCmd.Echo Not isOn
    DoCmd.SetWarnings Not isOn
    
    ' Excel側の制御
    If Not targetExcelApp Is Nothing Then
        On Error Resume Next
        targetExcelApp.ScreenUpdating = Not isOn
        targetExcelApp.EnableEvents = Not isOn
        If targetExcelApp.Workbooks.count > 0 Then
            targetExcelApp.Calculation = IIf(isOn, -4135, -4105) ' xlManual / xlAutomatic
        End If
        On Error GoTo 0
    End If
End Sub
' ← プロシージャの終わり
'--------------------------------------------
' 関数名： Get_FileName_EZ
' 概要： フルパスからファイル名部分（拡張子含む）を抽出する
'--------------------------------------------
Public Function Get_FileName_EZ(ByVal fullPath As String) As String
    If InStr(fullPath, "\") > 0 Then
        Get_FileName_EZ = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    Else
        Get_FileName_EZ = fullPath
    End If
End Function

'--------------------------------------------
' プロシージャ名： Sync_Registry_Path
' 概要： _at_SystemRegistry の既定パスを更新する（学習機能）
'--------------------------------------------
Public Sub Sync_Registry_Path(ByVal taskID As Long, ByVal newPath As String)
    If taskID <= 0 Or newPath = "" Then Exit Sub
    
    Dim db As DAO.Database: Set db = CurrentDb
    Dim sql As String
    
    ' フォルダパスの場合、末尾の \ を補完して正規化
    If Right(newPath, 1) <> "\" And InStr(Get_FileName_EZ(newPath), ".") = 0 Then
        newPath = newPath & "\"
    End If
    
    sql = "UPDATE [_at_SystemRegistry] SET [既定パス] = '" & Replace(newPath, "'", "''") & "' WHERE [ID] = " & taskID
    db.Execute sql, dbFailOnError
End Sub

'--------------------------------------------
' 関数名： G_GetSheetByCodeName
' 概要： 指定したExcelブックの中から、CodeName（オブジェクト名）またはシート名を基にシートを特定する
'--------------------------------------------
Public Function G_GetSheetByCodeName(ByRef wb As Object, ByVal nameOrCode As String) As Object
    Dim sh As Object
    ' 1. オブジェクト名（CodeName）で検索
    For Each sh In wb.Sheets
        If sh.codeName = nameOrCode Then
            Set G_GetSheetByCodeName = sh
            Exit Function
        End If
    Next sh
    
    ' 2. 見つからない場合はシート名（見出し名）で検索
    On Error Resume Next
    Set G_GetSheetByCodeName = wb.Sheets(nameOrCode)
    On Error GoTo 0
End Function

'--------------------------------------------
' 関数名： Fetch_New_ImportID
' 概要： 重複しない一意な数字IDを生成する (Access Long型 21億の制限内)
' 設計： 月(2桁) + 日(2桁) + 連番(6桁) で構成
'--------------------------------------------
Public Function Fetch_New_ImportID() As Long
    Static s_counter As Long
    Dim basePart As Long
    
    ' 月・日 をベースに作成 (例: 4月9日 -> 0409000000)
    basePart = CLng(Month(Now)) * 100 + CLng(Day(Now))
    
    ' セッション内連番をインクリメント
    If s_counter = 0 Then
        ' 初回のみ、現在の時分秒をシード値に近い形で足すか、シンプルに1から開始
        s_counter = 1
    Else
        s_counter = s_counter + 1
    End If
    
    ' 合体させて 10桁の数字にする (最大 1231 + 999999)
    ' Long型の最大 2,147,483,647 を超えないよう、ベースを100万倍にする
    ' 1231 * 1,000,000 + 999,999 = 1,231,999,999 (OK)
    Fetch_New_ImportID = (basePart * 1000000) + s_counter
End Function
