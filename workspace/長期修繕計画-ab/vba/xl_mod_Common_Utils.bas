Attribute VB_Name = "xl_mod_Common_Utils"

'-------------------------------------
' Module: xl_mod_Common_Utils
' 説明  ：システム全体で使用する汎用部品（ユーティリティ） '[cite: 16]
' 更新日：2026/01/10
'-------------------------------------
Option Explicit

' 一括実行中かどうかを管理するパブリック変数 '[cite: 16]
Public pIsBatchActive As Integer ' 0:単独実行, 1:一括実行中

'--------------------------------------------
' 関数名 : Normalize_Text
' 概要   : 文字列から全角・半角スペースを除去し、英数字を半角に統一する '[cite: 16]
'--------------------------------------------
Public Function Normalize_Text(ByVal txt As String) As String
    Dim result As String
    ' 全角・半角スペースをすべて除去 '[cite: 16]
    result = Replace(Replace(txt, " ", ""), "　", "")
    ' 英数字・記号を半角に統一 '[cite: 16]
    On Error Resume Next
    result = StrConv(result, vbNarrow)
    On Error GoTo 0
    Normalize_Text = Trim(result) '[cite: 17]
End Function

'--------------------------------------------
' 関数名 : Get_ColumnIndex_Robust
' 概要   : テーブルのタイトル名から列番号を取得する（表記ゆれ対応） '[cite: 18]
'--------------------------------------------
Public Function Get_ColumnIndex_Robust(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim col As ListColumn
    Dim target As String: target = Normalize_Text(colName) '[cite: 18]
    
    For Each col In lo.ListColumns
        ' テーブル側の見出しも正規化して比較 '[cite: 18]
        If Normalize_Text(col.Name) = target Then
            Get_ColumnIndex_Robust = col.Index
            Exit Function
        End If
    Next col
    Get_ColumnIndex_Robust = 0 '[cite: 19]
End Function

'--------------------------------------------
' プロシージャ名 : Fast_Mode_Toggle
' 概要           : 高速実行モードのON/OFF（画面更新停止等） '[cite: 20]
'--------------------------------------------
Public Sub Fast_Mode_Toggle(ByVal isOn As Boolean)
    With Application
        .ScreenUpdating = Not isOn
        .Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not isOn
    End With
End Sub

'--------------------------------------------
' プロシージャ名 : Notify_Smart_Popup
' 概要           : 一括実行中でない場合のみ、自動閉鎖ポップアップを表示する '[cite: 19]
'--------------------------------------------
Public Sub Notify_Smart_Popup(ByVal msg As String, ByVal title As String, Optional ByVal iconType As Integer = 64)
    ' 一括実行フラグが立っている場合は表示しない '[cite: 20]
    If pIsBatchActive = 1 Then Exit Sub
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    ' 1秒後に自動で閉じる '[cite: 20]
    wsh.Popup msg, 1, title, iconType
End Sub
' ← プロシージャの終わり '[cite: 12]

