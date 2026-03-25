VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "com_clsAccessFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: com_clsAccessFetcher
' 説明：外部AccessファイルにADOで接続し、データを配列やRecordsetとして取得する。
'       参照設定に依存しないレイトバインディング方式を採用。
' 更新日：2026/03/24（MyLauncher統合・堅牢ADOアクセス版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pFilePath As String
Private pErr      As com_clsErrorUtility

'===========================================================
' サブルーチン名 : Init
' 引数 : filePath - 対象のAccessファイルフルパス
'        errorUtil - (任意) com_clsErrorUtility
' 機能 : 接続先とエラーハンドラを初期化する。
'===========================================================
Public Sub Init(ByVal filePath As String, Optional ByRef errorUtil As com_clsErrorUtility)
    pFilePath = filePath
    
    If errorUtil Is Nothing Then
        Set pErr = New com_clsErrorUtility
        pErr.Init False
    Else
        Set pErr = errorUtil
    End If
End Sub

'===========================================================
' 関数名 : FetchArray
' 引数 : sql - 実行するSELECT文、またはテーブル/クエリ名
' 戻り値 : 2次元配列 (Variant)
' 機能 : 結果セットを2次元配列として一括取得する。
'===========================================================
Public Function FetchArray(ByVal sql As String) As Variant
    On Error GoTo Err_Handler
    
    Dim rs As Object: Set rs = FetchRecordset(sql)
    If rs Is Nothing Then
        FetchArray = Empty
        Exit Function
    End If
    
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        FetchArray = Empty
        Exit Function
    End If

    ' GetRowsによる高速配列化
    ' 転置が必要になる点に注意（列, 行 の順で返る）
    FetchArray = rs.GetRows()
    
    rs.Close: Set rs = Nothing
    Exit Function

Err_Handler:
    pErr.Notify_Smart_Popup "データ取得に失敗しました: " & Err.description, "Fetch Error", vbCritical
    FetchArray = Empty
End Function

'===========================================================
' 関数名 : FetchRecordset
' 引数 : sql - SQL文またはオブジェクト名
' 戻り値 : ADODB.Recordset (Object)
' 機能 : ADO接続を確立し、レコードセットを返す。
'===========================================================
Public Function FetchRecordset(ByVal sql As String) As Object
    On Error GoTo Err_Handler
    
    Dim conn As Object: Set conn = CreateObject("ADODB.Connection")
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
    
    ' ACEプロバイダを使用して接続
    Dim connStr As String
    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pFilePath & ";"
    
    conn.Open connStr
    
    ' SQLか単一のテーブル名かを簡易判定
    Dim query As String
    If InStr(1, sql, "SELECT", vbTextCompare) > 0 Then
        query = sql
    Else
        query = "SELECT * FROM [" & sql & "]"
    End If
    
    ' 読み取り専用・前方スクロールのみで高速オープン
    rs.Open query, conn, 0, 1 ' 0=adOpenForwardOnly, 1=adLockReadOnly
    
    Set FetchRecordset = rs
    Exit Function

Err_Handler:
    pErr.Notify_Smart_Popup "接続エラー: " & Err.description, "DB Connection Error", vbCritical
    Set FetchRecordset = Nothing
End Function
