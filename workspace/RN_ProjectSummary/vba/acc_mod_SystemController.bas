Attribute VB_Name = "acc_mod_SystemController"
Option Compare Database
Option Explicit

'----------------------------------------------------------------
' 関数名 : ExecuteTaskByID
' 概要   : _at_SystemRegistry の ID に基づき処理を実行する
'----------------------------------------------------------------
Public Sub ExecuteTaskByID(ByVal taskID As Long)
    Dim db      As DAO.Database: Set db = CurrentDb
    Dim rs      As DAO.Recordset
    Dim strSQL  As String
    Dim clsLog  As New com_clsErrorUtility
    
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
    clsLog.Init isBatch:=True
    
    ' 3. ファイル種別に応じた実行分岐
    Select Case rs!ファイル種別
        Case "Access(自)"
            ' 自ファイルのマクロを実行
            Application.Run rs!実行マクロ名
            
        Case "Access(クエリ)"
            ' クエリをデータシートビューで開く
            DoCmd.OpenQuery rs!実行マクロ名, acViewNormal, acReadOnly
            
        Case "Excel"
            ' 外部Excelのマクロを実行 (共通エンジンを別途呼び出し)
            Call RunExternalExcelMacro(rs!既定パス, rs!実行マクロ名)
            
        Case "Access(外部)"
            ' 外部Accessのマクロを実行 (共通エンジンを別途呼び出し)
            ' Call RunExternalAccessMacro(...)
    End Select
    
    ' 4. 結果の更新
    UpdateLastStatus taskID, "成功"
    
    ' ファイル種別が「Access(クエリ)」以外の場合のみメッセージを出す
    If rs!ファイル種別 <> "Access(クエリ)" Then
        MsgBox rs!処理名称 & " が完了しました。", vbInformation
    End If
    
    Exit Sub

Err_Handler:
    UpdateLastStatus taskID, "エラー: " & Err.Description
    clsLog.Notify_Smart_Popup "処理実行エラー: " & rs!処理名称 & vbCrLf & Err.Description
End Sub

' 最終実行状況の更新用サブ
Private Sub UpdateLastStatus(ByVal taskID As Long, ByVal status As String)
    Dim strSQL As String
    strSQL = "UPDATE [_at_SystemRegistry] SET [最終実行日時] = Now(), [最終実行結果] = '" & status & "' WHERE ID = " & taskID
    CurrentDb.Execute strSQL, dbFailOnError
End Sub

'----------------------------------------------------------------
' 概要 : 外部Excelのマクロを実行する（独立インスタンス方式）
'----------------------------------------------------------------
Public Sub RunExternalExcelMacro(ByVal filePath As String, ByVal macroName As String)
    Dim xlApp As Object
    Dim wb    As Object
    
    On Error GoTo Err_Handle
    
    ' 新しいExcelインスタンスを生成（既存のExcelとは無関係に動作）
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' 処理中は非表示
    
    ' ファイルを開く
    Set wb = xlApp.Workbooks.Open(filePath)
    
    ' マクロ実行
    xlApp.Run macroName
    
    ' 保存して閉じる
    wb.Close saveChanges:=True
    
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

'----------------------------------------------------------------
' 概要 : 外部Accessのマクロを実行する
'----------------------------------------------------------------
Public Sub RunExternalAccessMacro(ByVal filePath As String, ByVal procName As String)
    Dim acApp As Object
    
    On Error GoTo Err_Handle
    
    Set acApp = CreateObject("Access.Application")
    acApp.OpenCurrentDatabase filePath
    
    ' 公開されたプロシージャを実行
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
