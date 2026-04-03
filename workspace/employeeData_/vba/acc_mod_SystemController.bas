Attribute VB_Name = "acc_mod_SystemController"
Option Compare Database
Option Explicit

'----------------------------------------------------------------
' Module: acc_mod_SystemController
' 説明   : _at_SystemRegistry の ID に基づき処理を実行する汎用コントローラー
' 参照先 : RN_ProjectSummary より移植
'----------------------------------------------------------------

'----------------------------------------------------------------
' 関数名 : ExecuteTaskByID
' 概要   : 指定されたIDのタスクを実行する
'----------------------------------------------------------------
Public Sub ExecuteTaskByID(ByVal taskID As Long)
    Dim db      As DAO.Database: Set db = CurrentDb
    Dim rs      As DAO.Recordset
    Dim strSQL  As String
    
    ' 1. レジストリ情報の取得
    strSQL = "SELECT * FROM [_at_SystemRegistry] WHERE [ID] = " & taskID
    Set rs = db.OpenRecordset(strSQL, dbOpenForwardOnly)
    
    If rs.EOF Then
        MsgBox "指定された処理IDが見つかりません: " & taskID, vbCritical
        Exit Sub
    End If
    
    ' 2. 実行前確認
    If Nz(rs!実行前確認メッセージ, "") <> "" Then
        If MsgBox(rs!実行前確認メッセージ, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    End If
    
    On Error GoTo Err_Handler
    
    ' 3. ファイル種別に応じた実行分岐
    Select Case rs!ファイル種別
        Case "Access(自)"
            ' 自ファイルのマクロ/関数を実行 (引数として ID を渡す)
            If Nz(rs!実行マクロ名, "") <> "" Then
                Application.Run rs!実行マクロ名, taskID
            End If
            
        Case "Access(クエリ)"
            ' クエリを開く
            DoCmd.OpenQuery rs!実行マクロ名, acViewNormal, acReadOnly
            
        Case "Excel"
            ' 外部Excelのマクロを実行
            Call RunExternalExcelMacro(rs!既定パス, rs!実行マクロ名)
            
        Case "Access(外部)"
            ' 外部Accessのマクロを実行
            Call RunExternalAccessMacro(rs!既定パス, rs!実行マクロ名)
    End Select
    
    ' 4. 結果の更新
    UpdateLastStatus taskID, "成功"
    
    ' メッセージ表示（クエリ以外）
    If rs!ファイル種別 <> "Access(クエリ)" Then
        MsgBox rs!処理名称 & " が完了しました。", vbInformation
    End If
    
    Exit Sub

Err_Handler:
    UpdateLastStatus taskID, "エラー: " & Err.Description
    MsgBox "処理実行エラー: " & rs!処理名称 & vbCrLf & Err.Description, vbCritical
End Sub

' 最終実行状況の更新
Private Sub UpdateLastStatus(ByVal taskID As Long, ByVal status As String)
    Dim strSQL As String
    strSQL = "UPDATE [_at_SystemRegistry] SET [最終実行日時] = Now(), [最終実行結果] = '" & status & "' WHERE ID = " & taskID
    CurrentDb.Execute strSQL, dbFailOnError
End Sub

' 外部Excelマクロ実行
Public Sub RunExternalExcelMacro(ByVal filePath As String, ByVal macroName As String)
    Dim xlApp As Object
    Dim wb    As Object
    
    On Error GoTo Err_Handle
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set wb = xlApp.Workbooks.Open(filePath)
    xlApp.Run macroName
    wb.Close SaveChanges:=True
    
Clean_Up:
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    Exit Sub

Err_Handle:
    Err.Raise Err.Number, "RunExternalExcelMacro", "Excel実行失敗: " & Err.Description
    Resume Clean_Up
End Sub

' 外部Accessマクロ実行
Public Sub RunExternalAccessMacro(ByVal filePath As String, ByVal procName As String)
    Dim acApp As Object
    On Error GoTo Err_Handle
    Set acApp = CreateObject("Access.Application")
    acApp.OpenCurrentDatabase filePath
    acApp.Run procName
    acApp.CloseCurrentDatabase
    
Clean_Up:
    If Not acApp Is Nothing Then
        acApp.Quit
        Set acApp = Nothing
    End If
    Exit Sub

Err_Handle:
    Err.Raise Err.Number, "RunExternalAccessMacro", "Access実行失敗: " & Err.Description
    Resume Clean_Up
End Sub
