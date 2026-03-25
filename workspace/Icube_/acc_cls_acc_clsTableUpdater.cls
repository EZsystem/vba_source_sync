VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsTableUpdater
' 説明：DAO.Recordsetを使用してテーブルのフィールド値を一括・条件付更新する。
'       DIにより接続と通知を管理し、Where条件によるフィルタリングに対応。
' 更新日：2026/03/24（MyLauncher統合・高度な更新ロジック対応版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pConn          As acc_clsDbConnector
Private pErr           As com_clsErrorUtility
Private pTableName     As String
Private pWhereCondition As String

'===========================================================
' サブルーチン名 : Init
' 引数 : connector - acc_clsDbConnector インスタンス
'        tableName - 更新対象テーブル名
'        errorUtil - (任意) com_clsErrorUtility
' 機能 : 依存オブジェクトと対象テーブルを初期化する。
'===========================================================
Public Sub Init(ByRef connector As acc_clsDbConnector, _
                ByVal TableName As String, _
                Optional ByRef errorUtil As com_clsErrorUtility)
    Set pConn = connector
    pTableName = TableName
    
    If errorUtil Is Nothing Then
        Set pErr = New com_clsErrorUtility
        pErr.Init False
    Else
        Set pErr = errorUtil
    End If
End Sub

' --- プロパティ ---
Public Property Let WhereCondition(ByVal val As String): pWhereCondition = val: End Property
Public Property Get WhereCondition() As String: WhereCondition = pWhereCondition: End Property

'===========================================================
' サブルーチン名 : UpdateFieldFromField
' 引数 : srcFieldName - 値のコピー元フィールド
'        tgtFieldName - 更新先のフィールド
' 機能 : src の値を tgt に全件（または条件合致分）コピーする。
'===========================================================
Public Sub UpdateFieldFromField(ByVal srcFieldName As String, ByVal tgtFieldName As String)
    Dim rs As DAO.Recordset
    Set rs = OpenRecordsetSafe()
    If rs Is Nothing Then Exit Sub
    
    On Error GoTo Err_Handler
    Do Until rs.EOF
        rs.Edit
        rs.Fields(tgtFieldName).Value = rs.Fields(srcFieldName).Value
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub

Err_Handler:
    HandleError "UpdateFieldFromField", rs
End Sub

'===========================================================
' サブルーチン名 : UpdateFieldByConcat
' 引数 : fldA, fldB - 結合対象のフィールド
'        tgtField   - 結果格納先
'        separator  - 区切り文字 (既定: "_")
' 機能 : 2つのフィールドを結合して格納する。
'===========================================================
Public Sub UpdateFieldByConcat(ByVal fldA As String, ByVal fldB As String, _
                               ByVal tgtField As String, _
                               Optional ByVal separator As String = "_")
    Dim rs As DAO.Recordset
    Set rs = OpenRecordsetSafe()
    If rs Is Nothing Then Exit Sub
    
    On Error GoTo Err_Handler
    Do Until rs.EOF
        rs.Edit
        rs.Fields(tgtField).Value = Nz(rs.Fields(fldA).Value, "") & separator & Nz(rs.Fields(fldB).Value, "")
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub

Err_Handler:
    HandleError "UpdateFieldByConcat", rs
End Sub

'===========================================================
' 内部関数 : OpenRecordsetSafe
' 機能 : 条件設定に基づき、安全にレコードセットをオープンする。
'===========================================================
Private Function OpenRecordsetSafe() As DAO.Recordset
    On Error GoTo Err_Handler
    Dim sql As String
    sql = "SELECT * FROM [" & pTableName & "]"
    If pWhereCondition <> "" Then
        sql = sql & " WHERE " & pWhereCondition
    End If
    
    Set OpenRecordsetSafe = pConn.db.OpenRecordset(sql, dbOpenDynaset)
    Exit Function

Err_Handler:
    pErr.Notify_Smart_Popup "Recordsetのオープンに失敗しました: " & Err.description, "Open Error", vbCritical
    Set OpenRecordsetSafe = Nothing
End Function

'===========================================================
' 内部サブルーチン : HandleError
'===========================================================
Private Sub HandleError(ByVal procName As String, ByRef rs As DAO.Recordset)
    pErr.Notify_Smart_Popup "更新中にエラーが発生しました (" & procName & "): " & Err.description, "Update Error", vbCritical
    If Not rs Is Nothing Then rs.Close
    ' トランザクション中断のため再スロー
    Err.Raise Err.Number, "acc_clsTableUpdater." & procName, Err.description
End Sub
