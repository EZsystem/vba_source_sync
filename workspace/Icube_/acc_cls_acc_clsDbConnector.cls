VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsDbConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsDbConnector
' 説明：DAO.Database接続のライフサイクルを管理するコアクラス。
'       CurrentDbおよび外部ACCDB/MDBへの接続を透過的に提供する。
'       旧 acc_clsDbManager の機能を統合した最終版。
' 更新日：2026/03/24（MyLauncher統合・Manager機能吸収版）
'===========================================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private pDB         As DAO.Database
Private pFilePath   As String
Private pIsExternal As Boolean

'===========================================================
' サブルーチン名 : Init
' 引数 : filePath - 外部DBパス（省略時は CurrentDb に接続）
' 機能 : データベース接続を初期化する。
'===========================================================
Public Sub Init(Optional ByVal filePath As String = "")
    On Error GoTo Err_Handler
    
    pFilePath = Trim(filePath)
    
    ' 既存の接続がある場合は一旦クローズ（再初期化対応）
    If pIsExternal And Not pDB Is Nothing Then
        pDB.Close
    End If
    
    If pFilePath = "" Then
        ' --- 内部（カレント）DBへの接続 ---
        Set pDB = CurrentDb
        pIsExternal = False
    Else
        ' --- 外部DBへの接続 ---
        Set pDB = DBEngine.OpenDatabase(pFilePath, False, False)
        pIsExternal = True
    End If
    
    Exit Sub

Err_Handler:
    Dim errDetail As String
    errDetail = "DB接続に失敗しました。" & vbCrLf & "Path: " & pFilePath & vbCrLf & "Error: " & Err.description
    Err.Raise vbObjectError + 1001, "acc_clsDbConnector.Init", errDetail
End Sub

'===========================================================
' プロパティ Get : db
' 機能 : 保持している DAO.Database オブジェクトを返す。
'===========================================================
Public Property Get db() As DAO.Database
    ' 未初期化の場合はカレントDBで自動初期化を試みる（ガードロジック）
    If pDB Is Nothing Then
        Me.Init
    End If
    Set db = pDB
End Property

'===========================================================
' プロパティ Get : IsExternal
' 機能 : 現在の接続が外部DBかどうかを返す。
'===========================================================
Public Property Get IsExternal() As Boolean
    IsExternal = pIsExternal
End Property

'===========================================================
' サブルーチン名 : ExecuteSQL
' 引数 : sql - 実行するアクションクエリ
'        failOnError - エラー時に中断するか（既定：True）
' 機能 : SQLを実行するラッパー。
'===========================================================
Public Sub ExecuteSQL(ByVal sql As String, Optional ByVal failOnError As Boolean = True)
    Dim options As Long
    If failOnError Then options = dbFailOnError
    
    Me.db.Execute sql, options
End Sub

'===========================================================
' 関数名 : OpenRecordsetSafe
' 引数 : sql - 抽出用SQL文、またはテーブル名
' 戻り値 : DAO.Recordset (Dynaset)
' 機能 : 安全にレコードセットをオープンして返す。
'===========================================================
Public Function OpenRecordsetSafe(ByVal sql As String) As DAO.Recordset
    On Error GoTo Err_Handler
    
    ' dbプロパティ経由で取得することで未初期化を防止
    Set OpenRecordsetSafe = Me.db.OpenRecordset(sql, dbOpenDynaset)
    
    Exit Function

Err_Handler:
    Dim errDetail As String
    errDetail = "レコードセットの取得に失敗しました。" & vbCrLf & "SQL: " & sql & vbCrLf & "Error: " & Err.description
    Err.Raise vbObjectError + 1002, "acc_clsDbConnector.OpenRecordsetSafe", errDetail
End Function

'===========================================================
' クリーンアップ : Class_Terminate
' 機能 : 外部接続の場合は確実にクローズしてファイルロックを解除する。
'===========================================================
Private Sub Class_Terminate()
    On Error Resume Next
    If pIsExternal Then
        If Not pDB Is Nothing Then
            pDB.Close
        End If
    End If
    Set pDB = Nothing
End Sub
