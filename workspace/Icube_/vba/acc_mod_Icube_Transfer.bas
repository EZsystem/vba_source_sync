Attribute VB_Name = "acc_mod_Icube_Transfer"
'Attribute VB_Name = "acc_mod_Icube_Transfer"
'===================================================================================================
' モジュール名   : acc_mod_Icube_Transfer
' 概要           : バリデーション済みデータの統合および関連テーブルへの動的転写
' 依存コンポーネント:
'   - クラス     : com_clsErrorUtility (共通エラー/ログ管理)
'   - ライブラリ : Microsoft DAO 3.6 Object Library
' 最終更新日     : 2026/03/26
'===================================================================================================

Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Transfer_Icube_To_History
' 概要           : 作業用テーブル(Icube_)から累積テーブル(Icube_累計)へデータを統合します。
' 処理内容       : 枝番工事コードが重複する既存データを削除した後、最新データを全件挿入します。
' 引数           : ErrorLog (com_clsErrorUtility) - エラー情報を集約するクラスインスタンス
'---------------------------------------------------------------------------------------------------
Public Sub Transfer_Icube_To_History(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim DbObj As DAO.Database
    Dim SqlDelete As String
    Dim SqlInsert As String

    Set DbObj = CurrentDb
    
    ' 1. 重複削除：枝番工事コードが一致する既存データを累積側から削除（物理置換準備）
    SqlDelete = "DELETE FROM Icube_累計 WHERE [枝番工事コード] IN (SELECT [枝番工事コード] FROM Icube_)"
    DbObj.Execute SqlDelete, dbFailOnError

    ' 2. 全件追加：検証・クレンジングが完了した最新レコードを挿入
    SqlInsert = "INSERT INTO Icube_累計 SELECT * FROM Icube_"
    DbObj.Execute SqlInsert, dbFailOnError
    
    Exit Sub

Err_Handler:
    ErrorLog.Notify_Smart_Popup "Transfer History Error: " & Err.Description, "Transfer Error", vbCritical
End Sub

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Transfer_To_Related_Tables
' 概要           : Icube_（当月分）および Icube_累計（過去分含）から、各関連マスタテーブルへ配分します。
' 処理内容       : 定義された関連テーブル（kt_XXX）に対し、キーの存在確認を行いながら転写を実行します。
' 引数           : ErrorLog (com_clsErrorUtility) - エラー情報を集約するクラスインスタンス
'---------------------------------------------------------------------------------------------------
Public Sub Transfer_To_Related_Tables(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    
    ' 転写元テーブルのリスト（当月ワークと累積の両方を対象とする）
    Dim SourceTables As Variant
    SourceTables = Array("Icube_", "Icube_累計")
    
    Dim i As Integer
    Dim SrcName As String
    
    For i = 0 To UBound(SourceTables)
        SrcName = SourceTables(i)
        
        ' 各関連テーブル（接頭辞 kt_）への転写実行。
        ' 汎用エンジンを使用することで、フィールド構成の変更に柔軟に対応
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

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : TransferTable_Generic (汎用転写エンジン)
' 概要           : キーの重複をチェックしながら、フィールド名が一致する値のみを自動で転写します。
' 引数           : SrcTable (String) - 転写元テーブル名
'                  DstTable (String) - 転写先テーブル名
'                  KeyField (String) - 重複判定に使用するキー列名
'                  ErrorLog (com_clsErrorUtility) - エラーハンドラ
'---------------------------------------------------------------------------------------------------
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
        ' キー値の取得と、SQLインジェクション/エラー防止のためのシングルクォートエスケープ
        KeyValue = Replace(Nz(RsSrc(KeyField).Value, ""), "'", "''")
        
        ' ターゲット側に同一キーが存在するか確認（未存在時のみ追加する「重複回避」ロジック）
        MatchCount = DCount("*", "[" & DstTable & "]", "[" & KeyField & "] = '" & KeyValue & "'")
        
        If MatchCount = 0 Then
            RsDst.AddNew
            
            ' ソース側の全フィールドを走査し、ターゲット側に存在する同名フィールドのみ値を代入
            For Each fld In RsSrc.fields
                If Internal_FieldExists(RsDst, fld.Name) Then
                    On Error Resume Next ' 個別フィールドの型不一致等による中断を防止
                    RsDst(fld.Name).Value = fld.Value
                    
                    If Err.Number <> 0 Then
                        ' デバッグログへの記録に留め、全体の処理は継続
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

'---------------------------------------------------------------------------------------------------
' 関数名 : Internal_FieldExists
' 概要   : 指定したレコードセット内に、特定のフィールド名が存在するか判定します。
' 引数   : RsObj (DAO.Recordset) - 調査対象のレコードセット
'          FieldName (String) - 確認するフィールド名
' 戻り値 : Boolean - 存在すればTrue
'---------------------------------------------------------------------------------------------------
Private Function Internal_FieldExists(ByRef RsObj As DAO.Recordset, ByVal FieldName As String) As Boolean
    On Error Resume Next
    Dim FldObj As DAO.Field
    Set FldObj = RsObj.fields(FieldName)
    Internal_FieldExists = (Err.Number = 0)
    On Error GoTo 0
End Function
