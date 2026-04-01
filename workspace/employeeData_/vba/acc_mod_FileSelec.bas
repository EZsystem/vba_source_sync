Attribute VB_Name = "acc_mod_FileSelec"
'~~~~~~~~~~~~~~ acc_mod_FileSelec ~~~~~~~~~~~~~~
Option Compare Database
Option Explicit

'=================================================
' 関数名 : SelectExcelFileXLSX
' 概要   : 単一のExcelファイルを選択
'=================================================
Public Function SelectExcelFileXLSX() As String
    On Error GoTo ErrHandle
    Dim fd As FileDialog
    Dim selectedPath As String: selectedPath = ""
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker = 3
    With fd
        .title = "Excelファイル(xlsx)を選択してください"
        .Filters.Clear
        .Filters.Add "Excel(xlsx)", "*.xlsx"
        .AllowMultiSelect = False
        If .Show = -1 Then selectedPath = .SelectedItems(1)
    End With
    Set fd = Nothing
    SelectExcelFileXLSX = selectedPath
    Exit Function
ErrHandle:
    MsgBox "ファイル選択中にエラー発生:" & vbCrLf & Err.Description, vbExclamation
    SelectExcelFileXLSX = ""
End Function

'=================================================
' 関数名 : SelectFolder
' 概要   : フォルダを選択し、そのフルパスを返す
' 引数   : initialPath - 初期表示フォルダパス
'=================================================
Public Function SelectFolder(Optional ByVal initialPath As String = "") As String
    On Error GoTo ErrHandle
    Dim fd As FileDialog
    Dim selectedPath As String: selectedPath = ""
    ' 4: msoFileDialogFolderPicker
    Set fd = Application.FileDialog(4)
    With fd
        .title = "インポート対象のフォルダを選択してください"
        If initialPath <> "" Then .InitialFileName = initialPath
        If .Show = -1 Then
            selectedPath = .SelectedItems(1)
            ' 末尾に \ を付与
            If Right(selectedPath, 1) <> "\" Then selectedPath = selectedPath & "\"
        End If
    End With
    Set fd = Nothing
    SelectFolder = selectedPath
    Exit Function
ErrHandle:
    MsgBox "フォルダ選択中にエラー発生:" & vbCrLf & Err.Description, vbExclamation
    SelectFolder = ""
End Function

'=================================================
' 関数名 : SelectMultipleFiles
' 概要   : 複数のExcelファイルを選択し、コレクションを返す
' 引数   : initialPath - 初期表示フォルダパス
'=================================================
Public Function SelectMultipleFiles(Optional ByVal initialPath As String = "") As Collection
    On Error GoTo ErrHandle
    Dim fd As FileDialog
    Dim selectedFiles As New Collection
    Dim i As Long
    ' 3: msoFileDialogFilePicker
    Set fd = Application.FileDialog(3)
    With fd
        .title = "インポート対象のExcelファイルを選択してください（複数可）"
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xls*"
        .AllowMultiSelect = True
        If initialPath <> "" Then .InitialFileName = initialPath
        
        If .Show = -1 Then
            For i = 1 To .SelectedItems.Count
                selectedFiles.Add .SelectedItems(i)
            Next i
        End If
    End With
    Set fd = Nothing
    Set SelectMultipleFiles = selectedFiles
    Exit Function
ErrHandle:
    MsgBox "ファイル選択中にエラー発生:" & vbCrLf & Err.Description, vbExclamation
    Set SelectMultipleFiles = New Collection
End Function

