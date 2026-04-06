Attribute VB_Name = "sheet11_in"
Option Explicit

Sub sheet11_inFile()
    ' このコードは、ユーザーが指定したExcelファイルからデータを転写します。
    ' 取り込み元ファイルのデータを、取り込み先ファイルのシート「入力補助」に転写します。
    ' 取り込み前に、取り込み先のデータをクリアします。
    
    Dim sourceFile As String
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' ファイルダイアログを開いて、ユーザーにファイルを選択させる
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx"
        .title = "転写元のファイルを選択してください"
        If .Show = -1 Then
            sourceFile = .SelectedItems(1)
        Else
            MsgBox "ファイルが選択されませんでした。", vbExclamation
            Exit Sub
        End If
    End With
    
    ' 画面更新と自動計算を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' 取り込み先のシートを設定
    Set targetSheet = ThisWorkbook.Sheets("I22_Icube加工ALL")
    
    ' 取り込み先のデータをクリア
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).row
    lastCol = targetSheet.Cells(6, targetSheet.Columns.Count).End(xlToLeft).Column
    If lastRow >= 6 And lastCol >= 1 Then
        targetSheet.Range(targetSheet.Cells(6, 1), targetSheet.Cells(lastRow, lastCol)).ClearContents
    End If
    
    ' 取り込み元のファイルを開く
    Set sourceWorkbook = Workbooks.Open(sourceFile)
    Set sourceSheet = sourceWorkbook.Sheets(1)
    
    ' 取り込み元のデータ範囲を取得
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).row
    lastCol = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
    
    ' 取り込み元のデータを取り込み先に転写
    sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastRow, lastCol)).Copy
    targetSheet.Cells(6, 1).PasteSpecial Paste:=xlPasteValues
    targetSheet.Cells(6, 1).PasteSpecial Paste:=xlPasteFormats
    
    ' 取り込み元のファイルを閉じる
    sourceWorkbook.Close SaveChanges:=False
    
    ' 画面更新と自動計算を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    ' 完了メッセージ
    MsgBox "データの転写が完了しました。", vbInformation
End Sub
