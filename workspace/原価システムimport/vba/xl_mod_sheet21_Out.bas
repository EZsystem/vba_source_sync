Attribute VB_Name = "sheet21_Out"
Option Explicit

Sub sheet21_Out_1()
    '概要: 対象シートの指定範囲を別ファイルにコピーする自動化スクリプト
    'シート名: S1_受注、完工、既払い

    Dim ws As Worksheet
    Dim newBook As Workbook
    Dim newSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dateTimeStr As String
    Dim fileName As String
    Dim folderPath As String

    ' シートを設定
    Set ws = ThisWorkbook.Sheets("S1_受注、完工、既払い")
    
    ' パフォーマンス最適化の設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' 最終行と最終列を特定
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    
    ' 新しいブックを作成
    Set newBook = Workbooks.Add
    Set newSheet = newBook.Sheets(1)
    
    ' コピー範囲を設定
    ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol)).Copy
    
    ' 新しいシートに貼り付け
    With newSheet
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("A1").PasteSpecial Paste:=xlPasteFormats
        .Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    End With
    
    ' 日付と時間を取得
    dateTimeStr = Format(Now, "yyyymmdd_hhmmss")
    
    ' ファイル名を設定
    fileName = "受注、完工、既払い_" & dateTimeStr & ".xlsx"
    
    ' 保存フォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "保存先のフォルダを選択してください"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "フォルダが選択されませんでした。処理を中止します。", vbExclamation
            Exit Sub
        End If
    End With
    
    ' ファイルを保存
    newBook.SaveAs fileName:=folderPath & "\" & fileName, FileFormat:=xlOpenXMLWorkbook
    
    ' 新しいブックを閉じる
    newBook.Close SaveChanges:=False
    
    ' パフォーマンス設定を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' クリップボードをクリア
    Application.CutCopyMode = False
End Sub
