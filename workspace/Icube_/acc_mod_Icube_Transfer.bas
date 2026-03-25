Attribute VB_Name = "acc_mod_Icube_Transfer"
'Attribute VB_Name = "acc_mod_Icube_Transfer"
'===================================================================================================
' モジュール名 : acc_mod_Icube_Transfer
' 概要         : バりデーション済みデータの統合および関連テーブルへの転写
' 依存関係     : com_clsErrorUtility
'===================================================================================================

Option Compare Database
Option Explicit

'===========================================================
' プロシージャ名 : Transfer_Icube_To_History
' 概要           : 作業用テーブル(Icube_)から累積テーブル(Icube_累計)へデータを統合する
'===========================================================
Public Sub Transfer_Icube_To_History(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim DbObj As DAO.Database
    Dim SqlDelete As String
    Dim SqlInsert As String

    Set DbObj = CurrentDb
    
    ' 重複削除：枝番工事コードが一致する既存データを削除
    SqlDelete = "DELETE FROM Icube_累計 WHERE [枝番工事コード] IN (SELECT [枝番工事コード] FROM Icube_)"
    DbObj.Execute SqlDelete, dbFailOnError

    ' 全件追加：クレンジング済みのデータを挿入
    SqlInsert = "INSERT INTO Icube_累計 SELECT * FROM Icube_"
    DbObj.Execute SqlInsert, dbFailOnError
    
    Exit Sub

Err_Handler:
    ErrorLog.Notify_Smart_Popup "Transfer History Error: " & Err.Description, "Transfer Error", vbCritical
End Sub

'===========================================================
' プロシージャ名 : Transfer_To_Related_Tables
' 概要           : Icube_ および Icube_累計 から各 kt_ テーブルへ配分する
'===========================================================
Public Sub Transfer_To_Related_Tables(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    
    ' 転写元テーブルのリスト
    Dim SourceTables As Variant
    SourceTables = Array("Icube_", "Icube_累計")
    
    Dim i As Integer
    Dim SrcName As String
    
    For i = 0 To UBound(SourceTables)
        SrcName = SourceTables(i)
        
        ' 各関連テーブルへの転写実行
        Call TransferTable_Generic(SrcName, "kt_基本工事_完工", "基本工事コード", ErrorLog)
        Call TransferTable_Generic(SrcName, "kt_基本工事_作業所", "基本工事コード", ErrorLog)
        Call TransferTable_Generic(SrcName, "kt_基本工事_受注", "基本工事コード", ErrorLog)
        Call TransferTable_Generic(SrcName, "kt_工事コード情報", "工事コード", ErrorLog)
        Call TransferTable_Generic(SrcName, "kt_枝番工事", "枝番工事コード", ErrorLog)
    Next i
    
    Exit Sub

Err_Handler:
    ErrorLog.Notify_Smart_Popup "Related Tables Distribution Error: " & Err.Description
End Sub

'===========================================================
' プロシージャ名 : TransferTable_Generic (汎用転写エンジン)
' 概要           : キーの重複をチェックしながら、フィールド名が一致する値を転写する
'===========================================================
Private Sub TransferTable_Generic(ByVal SrcTable As String, ByVal DstTable As String, ByVal KeyField As String, ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim DbObj As DAO.Database
    Dim RsSrc As DAO.Recordset
    Dim RsDst As DAO.Recordset
    Dim fld As DAO.Field
    Dim KeyValue As String
    Dim MatchCount As Long

    Set DbObj = CurrentDb
    Set RsSrc = DbObj.OpenRecordset("SELECT * FROM [" & SrcTable & "]", dbOpenSnapshot)
    Set RsDst = DbObj.OpenRecordset(DstTable, dbOpenDynaset)

    Do While Not RsSrc.EOF
        ' キー値の取得とエスケープ処理
        KeyValue = Replace(Nz(RsSrc(KeyField).Value, ""), "'", "''")
        
        ' ターゲット側にキーが存在するか確認
        MatchCount = DCount("*", "[" & DstTable & "]", "[" & KeyField & "] = '" & KeyValue & "'")
        
        ' 未存在の場合のみ新規追加
        If MatchCount = 0 Then
            RsDst.AddNew
            
            ' ソース側の全フィールドをループ
            For Each fld In RsSrc.fields
                ' ターゲット側に同名のフィールドが存在するかチェック
                If Internal_FieldExists(RsDst, fld.Name) Then
                    On Error Resume Next
                    RsDst(fld.Name).Value = fld.Value
                    
                    ' 型不一致などのエラーハンドリング
                    If Err.Number <> 0 Then
                        Debug.Print "Field Transfer Skip: " & DstTable & "." & fld.Name & " (" & Err.Description & ")"
                        Err.Clear
                    End If
                    On Error GoTo Err_Handler
                End If
            Next fld
            
            RsDst.Update
        End If
        
        RsSrc.MoveNext
    Loop
    
    RsSrc.Close
    RsDst.Close
    Exit Sub

Err_Handler:
    ErrorLog.Notify_Smart_Popup DstTable & " Transfer Engine Error: " & Err.Description
End Sub

'===========================================================
' 関数名 : Internal_FieldExists
' 概要   : レコードセット内に指定した名前のフィールドがあるか確認する
'===========================================================
Private Function Internal_FieldExists(ByRef RsObj As DAO.Recordset, ByVal FieldName As String) As Boolean
    On Error Resume Next
    Dim FldObj As DAO.Field
    Set FldObj = RsObj.fields(FieldName)
    Internal_FieldExists = (Err.Number = 0)
    On Error GoTo 0
End Function

