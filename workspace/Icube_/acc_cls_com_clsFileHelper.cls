VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "com_clsFileHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: com_clsFileHelper
' 説明：ファイル・フォルダの存在確認、パス結合、テキスト入出力を管理する。
'       FileSystemObject (FSO) を内部で使用し、OSレベルの安全な操作を提供する。
' 更新日：2026/03/24（MyLauncher統合・堅牢パス管理版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pFso As Object

'===========================================================
' サブルーチン名 : Init
' 機能 : FileSystemObject を初期化する。
'===========================================================
Public Sub Init()
    Set pFso = CreateObject("Scripting.FileSystemObject")
End Sub

'===========================================================
' 関数名 : FileExists / FolderExists
' 引数 : path - 確認対象のパス
' 戻り値 : Boolean
'===========================================================
Public Function FileExists(ByVal path As String) As Boolean
    If pFso Is Nothing Then Me.Init
    FileExists = pFso.FileExists(path)
End Function

Public Function FolderExists(ByVal path As String) As Boolean
    If pFso Is Nothing Then Me.Init
    FolderExists = pFso.FolderExists(path)
End Function

'===========================================================
' 関数名 : SafePathJoin
' 引数 : baseDir - ベースとなるフォルダ
'        fileName - 結合するファイル名またはサブフォルダ名
' 戻り値 : 結合されたフルパス
' 機能 : フォルダ区切り文字（\）の有無を自動判定して安全に結合する。
'===========================================================
Public Function SafePathJoin(ByVal baseDir As String, ByVal fileName As String) As String
    If pFso Is Nothing Then Me.Init
    SafePathJoin = pFso.BuildPath(baseDir, fileName)
End Function

'===========================================================
' 関数名 : GetExtension
' 戻り値 : ファイルの拡張子（例: "xlsx"）
'===========================================================
Public Function GetExtension(ByVal path As String) As String
    If pFso Is Nothing Then Me.Init
    GetExtension = pFso.GetExtensionName(path)
End Function

'===========================================================
' サブルーチン名 : CreateFolderIfNotExist
' 引数 : path - 生成したいフォルダパス
' 機能 : 指定したフォルダが存在しない場合のみ、新規に作成する。
'===========================================================
Public Sub CreateFolderIfNotExist(ByVal path As String)
    If pFso Is Nothing Then Me.Init
    If Not pFso.FolderExists(path) Then
        pFso.CreateFolder path
    End If
End Sub

'===========================================================
' 関数名 : ReadTextFile
' 引数 : path - 読込対象
' 戻り値 : String (全テキスト内容)
' 機能 : 指定ファイルを一括で読み込む。
'===========================================================
Public Function ReadTextFile(ByVal path As String) As String
    If Not FileExists(path) Then Exit Function
    
    Dim ts As Object
    Set ts = pFso.OpenTextFile(path, 1) ' 1 = ForReading
    ReadTextFile = ts.ReadAll
    ts.Close
End Function

'===========================================================
' サブルーチン名 : WriteTextFile
' 引数 : path - 書き込み先
'        content - 内容
'        append - 追記モードにするか (Default: False)
'===========================================================
Public Sub WriteTextFile(ByVal path As String, ByVal content As String, Optional ByVal append As Boolean = False)
    If pFso Is Nothing Then Me.Init
    
    Dim mode As Integer: mode = IIf(append, 8, 2) ' 8=ForAppending, 2=ForWriting
    Dim ts As Object
    Set ts = pFso.OpenTextFile(path, mode, True) ' True = Create if not exist
    ts.Write content
    ts.Close
End Sub

'===========================================================
' クリーンアップ
'===========================================================
Private Sub Class_Terminate()
    Set pFso = Nothing
End Sub
