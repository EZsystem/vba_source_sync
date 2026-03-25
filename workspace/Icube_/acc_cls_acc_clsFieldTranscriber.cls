VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsFieldTranscriber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsFieldTranscriber
' 説明：テーブル内の任意フィールド間で値を転写するユーティリティ。
'       同一レコード内でのデータ移動や、Nullスキップ設定に対応。
' 更新日：2026/03/24（MyLauncher統合・DI対応版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pConn           As acc_clsDbConnector
Private pErr            As com_clsErrorUtility
Private pTableName      As String
Private pWhereCondition As String
Private pSkipNulls      As Boolean
Private pMappings       As Object ' Scripting.Dictionary (SourceField -> TargetField)

'===========================================================
' サブルーチン名 : Init
' 引数 : connector - acc_clsDbConnector インスタンス
'        tableName - 操作対象テーブル名
'        errorUtil - (任意) com_clsErrorUtility
' 機能 : 依存オブジェクトと対象テーブルを初期化する。
'===========================================================
Public Sub Init(ByRef connector As acc_clsDbConnector, _
                ByVal TableName As String, _
                Optional ByRef errorUtil As com_clsErrorUtility)
    Set pConn = connector
    pTableName = TableName
    pSkipNulls = True ' 既定でNullは転写しない（上書き防止）
    
    If errorUtil Is Nothing Then
        Set pErr = New com_clsErrorUtility
        pErr.Init False
    Else
        Set pErr = errorUtil
    End If
    
    Set pMappings = CreateObject("Scripting.Dictionary")
End Sub

' --- プロパティ設定 ---
Public Property Let WhereCondition(ByVal val As String): pWhereCondition = val: End Property
Public Property Let SkipNulls(ByVal val As Boolean): pSkipNulls = val: End Property

'===========================================================
' サブルーチン名 : AddMapping
' 引数 : srcFieldName - 転写元フィールド
'        tgtFieldName - 転写先フィールド
' 機能 : 転写ルール（元→先）を登録する。
'===========================================================
Public Sub AddMapping(ByVal srcFieldName As String, ByVal tgtFieldName As String)
    Dim s As String: s = Trim(srcFieldName)
    Dim t As String: t = Trim(tgtFieldName)
    If s <> "" And t <> "" Then
        pMappings(s) = t
    End If
End Sub

'===========================================================
' サブルーチン名 : TranscribeAll
' 機能 : 登録されたマッピングに基づき、レコードを走査して転写を実行する。
'===========================================================
Public Sub TranscribeAll()
    On Error GoTo Err_Handler
    
    If pMappings.Count = 0 Then Exit Sub
    
    Dim sql As String
    sql = "SELECT * FROM [" & pTableName & "]"
    If pWhereCondition <> "" Then sql = sql & " WHERE " & pWhereCondition
    
    Dim rs As DAO.Recordset
    Set rs = pConn.Db.OpenRecordset(sql, dbOpenDynaset)
    
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If
    
    Dim srcFld As Variant
    Dim val As Variant
    
    Do Until rs.EOF
        rs.Edit
        
        For Each srcFld In pMappings.Keys
            val = rs.fields(srcFld).Value
            
            ' SkipNulls設定に応じた処理
            If pSkipNulls And IsNull(val) Then
                ' スキップ（何もしない）
            Else
                rs.fields(pMappings(srcFld)).Value = val
            End If
        Next srcFld
        
        rs.Update
        rs.MoveNext
    Loop
    
    rs.Close
    pErr.Notify_Smart_Popup "フィールド間転写が完了しました: " & pTableName, "Transcribe Complete"
    Exit Sub

Err_Handler:
    pErr.Notify_Smart_Popup "転送中にエラーが発生しました: " & Err.Description, "Transcribe Error", vbCritical
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number
End Sub
