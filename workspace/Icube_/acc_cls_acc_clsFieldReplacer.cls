VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsFieldReplacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsFieldReplacer
' 説明：マスタ定義（Dictionary）に基づき、テーブル内の値を一括置換する。
'       特定のフィールドのみ、あるいは全フィールドを対象に検索置換が可能。
' 更新日：2026/03/24（MyLauncher統合・DI対応版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pConn         As acc_clsDbConnector
Private pErr          As com_clsErrorUtility
Private pTableName     As String
Private pDictMapping   As Object ' Scripting.Dictionary (置換前 -> 置換後)
Private pTargetFields  As Collection

'===========================================================
' サブルーチン名 : Init
' 引数 : connector - acc_clsDbConnector インスタンス
'        tableName - 置換対象のテーブル名
'        errorUtil - (任意) com_clsErrorUtility
' 機能 : 依存オブジェクトと対象テーブル、設定を初期化する。
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
    
    Set pDictMapping = CreateObject("Scripting.Dictionary")
    Set pTargetFields = New Collection
End Sub

'===========================================================
' サブルーチン名 : AddMapping
' 引数 : findValue - 置換前の値
'        replaceValue - 置換後の値
' 機能 : 置換ルールを登録する。
'===========================================================
Public Sub AddMapping(ByVal findValue As String, ByVal replaceValue As String)
    pDictMapping(findValue) = replaceValue
End Sub

'===========================================================
' サブルーチン名 : AddTargetField
' 引数 : fieldName - 置換対象とするフィールド名
' 機能 : 置換を行うフィールドを限定する場合に登録する。
'===========================================================
Public Sub AddTargetField(ByVal FieldName As String)
    pTargetFields.Add FieldName
End Sub

'===========================================================
' サブルーチン名 : ReplaceAll
' 機能 : 登録されたルールに基づき、レコードセットを走査して置換を実行する。
'===========================================================
Public Sub ReplaceAll()
    On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    Set rs = pConn.Db.OpenRecordset(pTableName, dbOpenDynaset)
    
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If

    Dim fld As DAO.Field
    Dim val As String
    
    Do Until rs.EOF
        rs.Edit
        
        ' 対象フィールドが指定されていればそれのみ、なければ全フィールドを走査
        If pTargetFields.Count > 0 Then
            Dim fldName As Variant
            For Each fldName In pTargetFields
                ProcessField rs.fields(fldName)
            Next fldName
        Else
            For Each fld In rs.fields
                ProcessField fld
            Next fld
        End If
        
        rs.Update
        rs.MoveNext
    Loop
    
    rs.Close
    pErr.Notify_Smart_Popup "一括置換が完了しました: " & pTableName, "Replace Complete"
    Exit Sub

Err_Handler:
    pErr.Notify_Smart_Popup "置換中にエラーが発生しました: " & Err.Description, "Replace Error", vbCritical
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number
End Sub

'===========================================================
' 内部サブルーチン : ProcessField
' 機能 : フィールドの値が辞書にあれば置換する。
'===========================================================
Private Sub ProcessField(ByRef fld As DAO.Field)
    Dim currentVal As String
    currentVal = Nz(fld.Value, "")
    
    If pDictMapping.Exists(currentVal) Then
        fld.Value = pDictMapping(currentVal)
    End If
End Sub
