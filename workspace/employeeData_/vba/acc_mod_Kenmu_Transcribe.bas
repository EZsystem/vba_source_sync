Attribute VB_Name = "acc_mod_Kenmu_Transcribe"
'-------------------------------------
' Module: acc_mod_Kenmu_Transcribe
' 概要  : 小数(0.5)・パーセント文字列(50%)・整数(50)を全て正しく0.5(50%)へ変換
' 修正内容: 「作業所名」の転写を追加
'-------------------------------------
Option Explicit

Private Const SRC_TABLE As String = "at_kenmuTemp"
Private Const TGT_TABLE As String = "at_kenmu"

Public Function Transcribe_Kenmu_Data_WithLog()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsSrc As DAO.Recordset
    Dim rsTgt As DAO.Recordset
    Dim trans As New acc_clsTransactionManager
    Dim errLog As New com_clsErrorUtility
    
    errLog.Initialize debugMode:=True
    Call Fast_Mode_Toggle(True)
    
    On Error GoTo ErrLine
    db.Execute "DELETE * FROM [" & TGT_TABLE & "];", dbFailOnError
    trans.BeginTransaction
    
    Set rsSrc = db.OpenRecordset(SRC_TABLE, dbOpenSnapshot)
    Set rsTgt = db.OpenRecordset(TGT_TABLE, dbOpenDynaset)
    
    Do Until rsSrc.EOF
        On Error Resume Next
        
        ' --- 兼務率の計算と判定 ---
        Dim dblRate As Double
        dblRate = Cleanse_Percent_Smart(rsSrc!兼務率割合, rsSrc!ImportID, errLog)
        
        ' ★ 0（またはエラーによる0）は除外する
        If dblRate <> 0 Then
            rsTgt.AddNew
            
            ' --- 【追加】作業所名の転記 ---
            rsTgt!作業所名 = rsSrc!作業所名
            
            ' 基本項目の転記
            rsTgt!ImportID = rsSrc!ImportID
            rsTgt!No = rsSrc!No
            rsTgt!工事コード = rsSrc!工事コード
            rsTgt!工事名 = rsSrc!工事名
            rsTgt!コメント = rsSrc!コメント
            rsTgt!社員名 = rsSrc!社員名
            ' ★パス情報を本テーブルへ引き継ぐ
            rsTgt!元ファイルパス = rsSrc!元ファイルパス
            rsTgt!作業所名 = rsSrc!作業所名
            
            ' 日付クレンジング（1日補正付）
            rsTgt!年月 = Cleanse_Date_Smart(rsSrc!年月, rsSrc!ImportID, errLog)
            
            ' 兼務率割合の転記 (0.5 等の小数として格納)
            rsTgt!兼務率割合 = dblRate
            
            ' エラーチェック
            If Err.Number <> 0 Then
                errLog.LogError "転記失敗(ID:" & rsSrc!ImportID & ")", Err.Description
                rsTgt.CancelUpdate
                Err.Clear
            Else
                rsTgt.Update
            End If
        End If
        
        rsSrc.MoveNext
        On Error GoTo ErrLine
    Loop
    
    ' 確定判定
    If errLog.ErrorCount > 0 Then
        If Application.UserControl Then
            ' 人間が操作している場合のみ確認を出す
            If MsgBox(errLog.ErrorCount & "件の異常がありました。ログを表示しますか？", vbYesNo) = vbYes Then
                trans.Rollback
                errLog.ShowAllErrors showInMsgBox:=True
                GoTo CleanUp
            End If
        Else
            ' Pythonからの実行時は、エラーがあれば即座に中断
            trans.Rollback
            GoTo CleanUp
        End If
    End If

    trans.Commit
    If Application.UserControl Then
        Call Notify_Smart_Popup("転写が完了しました（0%は除外済み）。", "完了")
    End If
CleanUp:
    If Not rsSrc Is Nothing Then rsSrc.Close
    If Not rsTgt Is Nothing Then rsTgt.Close
    Call Fast_Mode_Toggle(False)
    Exit Function

ErrLine:
    trans.Rollback
    If Application.UserControl Then
        MsgBox "致命的なエラー: " & Err.Description, vbCritical
    Else
        ' Python側でエラー内容を把握するため、イミディエイトウィンドウにだけ出す
        Debug.Print "Critical VBA Error: " & Err.Description
    End If
    Resume CleanUp
End Function


'--------------------------------------------
' 補助関数：兼務率のスマート変換（決定版）
' 判定ロジック:
' 1. "%"が含まれる場合(50%) -> 100で割る(0.5)
' 2. 1より大きい数値の場合(50) -> 100で割る(0.5)
' 3. 1以下の数値の場合(0.1) -> そのまま(0.1) ※10%として表示される
'--------------------------------------------
Private Function Cleanse_Percent_Smart(ByVal val As Variant, ByVal sourceID As Long, ByRef errLog As com_clsErrorUtility) As Double
    Dim sRaw As String: sRaw = Trim(Nz(val, ""))
    
    If sRaw = "" Or sRaw = "0" Then
        Cleanse_Percent_Smart = 0
        Exit Function
    End If
    
    ' 数値として扱えるかチェック
    Dim dVal As Double
    If IsNumeric(Replace(sRaw, "%", "")) Then
        dVal = CDbl(Replace(sRaw, "%", ""))
        
        ' パターン判定
        If InStr(sRaw, "%") > 0 Then
            ' "50%" -> 0.5
            Cleanse_Percent_Smart = dVal / 100
        ElseIf dVal > 1 Then
            ' "50" -> 0.5 (1より大きい場合は整数表記のパーセントとみなす)
            Cleanse_Percent_Smart = dVal / 100
        Else
            ' "0.5" -> 0.5 (1以下の場合は既に小数表記とみなす)
            Cleanse_Percent_Smart = dVal
        End If
    Else
        ' 数字ですらない場合
        errLog.LogError "数値形式異常", "ImportID: " & sourceID & " / 値: [" & sRaw & "]"
        Cleanse_Percent_Smart = 0
    End If
End Function

' 補助関数：日付変換（西暦・和暦両対応 / 1日補正）
Private Function Cleanse_Date_Smart(ByVal val As Variant, ByVal sourceID As Long, ByRef errLog As com_clsErrorUtility) As Variant
    Dim sRaw As String: sRaw = Trim(Nz(val, ""))
    If sRaw = "" Then Exit Function
    If IsDate(sRaw) Then
        Cleanse_Date_Smart = DateSerial(Year(CDate(sRaw)), Month(CDate(sRaw)), 1)
        Exit Function
    End If
    Dim sConv As String
    sConv = "H" & Replace(Replace(sRaw, "年", "/"), "月", "")
    If Right(sConv, 1) <> "/" Then sConv = sConv & "/1"
    If IsDate(sConv) Then
        Cleanse_Date_Smart = CDate(sConv)
    Else
        errLog.LogError "日付形式異常", "ImportID: " & sourceID & " / 値: [" & sRaw & "]"
        Cleanse_Date_Smart = Null
    End If
End Function
