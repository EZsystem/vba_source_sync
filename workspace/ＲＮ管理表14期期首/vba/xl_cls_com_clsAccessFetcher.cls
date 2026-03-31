VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "com_clsAccessFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: com_clsAccessFetcher
' 説明　：AccessファイルにADOで接続し、クエリやSQLの結果を配列またはRecordsetで返す共通クラスにゃ
' 作成日：2025/05/14
' 作成者：そうじろう
'=====================================
Option Explicit

' --- 内部保持変数 ---
Private pFilePath As String

'=================================================
' プロパティ Let/Get : FilePath
'=================================================
Public Property Let FilePath(val As String)
    pFilePath = val
End Property

Public Property Get FilePath() As String
    FilePath = pFilePath
End Property

'=================================================
' 関数名 : FetchArray
' 引数 : queryOrSql - クエリ名またはSQL文
' 戻り値 : 2次元配列（Variant）
'=================================================
Public Function FetchArray(queryOrSql As String) As Variant
    Dim rs As Object
    Set rs = FetchRecordset(queryOrSql)
    If rs Is Nothing Then Exit Function
    If rs.EOF Then rs.Close: Set rs = Nothing: Exit Function

    rs.MoveLast: rs.MoveFirst
    Dim arr() As Variant
    Dim i As Long, j As Long
    ReDim arr(1 To rs.RecordCount, 1 To rs.Fields.Count)
    Do Until rs.EOF
        i = i + 1
        For j = 1 To rs.Fields.Count
            arr(i, j) = rs.Fields(j - 1).value
        Next j
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    FetchArray = arr
End Function

'=================================================
' 関数名 : FetchFieldNames
' 引数 : queryOrSql - クエリ名またはSQL文
' 戻り値 : フィールド名の1次元配列（String型）
'=================================================
Public Function FetchFieldNames(queryOrSql As String) As String()
    Dim rs As Object
    Set rs = FetchRecordset(queryOrSql)
    If rs Is Nothing Then Exit Function

    Dim arr() As String
    Dim i As Long
    ReDim arr(1 To rs.Fields.Count)
    For i = 0 To rs.Fields.Count - 1
        arr(i + 1) = rs.Fields(i).Name
    Next i
    rs.Close
    Set rs = Nothing
    FetchFieldNames = arr
End Function

'=================================================
' 関数名 : FetchRecordset
' 引数 : queryOrSql - クエリ名またはSQL文
' 戻り値 : ADODB.Recordsetオブジェクト
'=================================================
Public Function FetchRecordset(queryOrSql As String) As Object
    Dim conn As Object, rs As Object
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    On Error GoTo ErrHandler
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pFilePath

    ' クエリ名 or SQL 判定（空白含むならSQLと判断）
    If InStr(queryOrSql, " ") > 0 Or InStr(queryOrSql, "SELECT") > 0 Then
        rs.Open queryOrSql, conn, 1, 1
    Else
        rs.Open "[" & queryOrSql & "]", conn, 1, 1
    End If

    Set FetchRecordset = rs
    Set conn = Nothing
    Exit Function

ErrHandler:
    MsgBox "FetchRecordsetエラー：" & Err.Description, vbCritical
    Set FetchRecordset = Nothing
    Set rs = Nothing: Set conn = Nothing
End Function

'=================================================
' クラス終了処理：不要なオブジェクト解放
'=================================================
Private Sub Class_Terminate()
    pFilePath = ""
End Sub

