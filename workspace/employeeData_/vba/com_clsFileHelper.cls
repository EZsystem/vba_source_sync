VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "com_clsFileHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: com_clsFileHelper
' 説明　：ファイル存在確認やファイル入出力操作を提供するユーティリティ
' 作成日：2025/04/25
' 更新日：再生成 by そうじろう
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
'（このクラスでは内部変数はありません）

'=================================================
' 関数名 : FileExists
' 引数 : path - 確認対象のファイルパス
' 機能 : 指定パスのファイルが存在するかどうかを返すにゃ
'=================================================
Public Function FileExists(path As String) As Boolean
    FileExists = (Dir(path) <> "")
End Function

'=================================================
' 関数名 : GetFileName
' 引数 : path - フルパス文字列
' 機能 : フルパスからファイル名部分だけを抽出して返すにゃ
'=================================================
Public Function GetFileName(path As String) As String
    GetFileName = Mid(path, InStrRev(path, "\") + 1)
End Function

'=================================================
' 関数名 : GetFolder
' 引数 : path - フルパス文字列
' 機能 : フルパスからフォルダパス部分を抽出して返すにゃ
'=================================================
Public Function GetFolder(path As String) As String
    GetFolder = Left(path, InStrRev(path, "\") - 1)
End Function

'=================================================
' 関数名 : ReadTextFile
' 引数 : path - 読込対象のファイルパス
' 戻り値 : ファイルの全内容（文字列）
' 機能 : 指定されたテキストファイルを開いて内容を一括読込するにゃ
'        存在しない場合はエラーを出すにゃ
'=================================================
Public Function ReadTextFile(path As String) As String
    Dim txt As String, f As Integer
    If Not FileExists(path) Then Err.Raise vbObjectError + 200, , "ファイルがありません"
    f = FreeFile
    Open path For Input As #f
    txt = Input$(LOF(f), f)
    Close #f
    ReadTextFile = txt
End Function

'=================================================
' サブルーチン名 : WriteTextFile
' 引数 : path - 出力先ファイルパス
'        content - 書き込み内容
' 機能 : 指定されたパスにテキストを書き込むにゃ（上書きモード）
'=================================================
Public Sub WriteTextFile(path As String, content As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, content
    Close #f
End Sub

'=================================================
' 関数名 : ShowFileDialog
' 引数 : title - ダイアログのタイトル
'        filter - フィルター設定 (例: "Excelファイル,*.xlsx;*.xlsm")
' 戻り値 : 選択されたファイルのフルパス（配列）
' 機能 : 標準のファイル選択ダイアログを表示するにゃ
'=================================================
Public Function ShowFileDialog(ByVal title As String, ByVal filter As String) As Variant
    Dim fd As Object
    ' AccessからOffice共通のFileDialogを呼び出す
    Set fd = Application.FileDialog(3) ' 3 = msoFileDialogFilePicker
    
    With fd
        .title = title
        .AllowMultiSelect = True
        .Filters.Clear
        
        ' フィルターの解析 (カンマ区切り想定)
        Dim parts() As String: parts = Split(filter, ",")
        If UBound(parts) >= 1 Then
            .Filters.Add parts(0), parts(1)
        End If
        
        If .Show = -1 Then
            ' 選択されたアイテムを配列に変換して返す
            Dim result() As Variant
            ReDim result(1 To .SelectedItems.Count)
            Dim i As Long
            For i = 1 To .SelectedItems.Count
                result(i) = .SelectedItems(i)
            Next i
            ShowFileDialog = result
        Else
            ShowFileDialog = Empty
        End If
    End With
End Function
