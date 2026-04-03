Attribute VB_Name = "mod_FileSelec"
'~~~~~~~~~~~~~~ mod_FileSelec ~~~~~~~~~~~~~~
Option Compare Database
Option Explicit

Public Function SelectExcelFileXLSX() As String
    On Error GoTo ErrHandle
    
    Dim fd As FileDialog
    Dim selectedPath As String
    selectedPath = ""

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Excelファイル(xlsx)を選択してください"
        .Filters.Clear
        ' 拡張子 xlsx のみ許可
        .Filters.Add "Excel(xlsx)", "*.xlsx"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            selectedPath = .SelectedItems(1)
        End If
    End With
    
    Set fd = Nothing
    SelectExcelFileXLSX = selectedPath
    
    Exit Function
    
ErrHandle:
    Debug.Print "SelectExcelFileXLSXでエラー発生: " & Err.Description
    MsgBox "エラー発生:" & vbCrLf & Err.Description, vbExclamation
    SelectExcelFileXLSX = ""
End Function

