Attribute VB_Name = "she031_In"
Option Explicit

' このコードはユーザーが指定したExcelファイルからデータを転写する
Sub she0321_In_1()
    ' 変数の定義
    Dim srcWorkbook As Workbook
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim srcFilename As String
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim destFirstRow As Long
    Dim destLastRow As Long
    Dim destFirstCol As Long
    Dim destLastCol As Long

    ' 初期設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' 転送先シートの設定
    Set destSheet = ThisWorkbook.Sheets("D1小口D")

    ' 取込み先のデータをクリアする
    firstRow = 7
    firstCol = 1
    destLastRow = destSheet.Cells(destSheet.Rows.Count, firstCol).End(xlUp).row
    destLastCol = destSheet.Cells(6, destSheet.Columns.Count).End(xlToLeft).Column

    If destLastRow >= firstRow And destLastCol >= firstCol Then
        destSheet.Range(destSheet.Cells(firstRow, firstCol), destSheet.Cells(destLastRow, destLastCol)).ClearContents
    End If

    ' 取込み元ファイルを指定
    srcFilename = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")
    If srcFilename = "False" Then
        MsgBox "ファイルが選択されませんでした。処理を終了します。"
        Exit Sub
    End If

    ' 取込み元ファイルを開く
    Set srcWorkbook = Workbooks.Open(srcFilename)
    Set srcSheet = srcWorkbook.Sheets(1)

    ' 取込み元範囲の設定
    firstRow = 2
    firstCol = 1
    lastRow = srcSheet.Cells(srcSheet.Rows.Count, firstCol).End(xlUp).row
    lastCol = srcSheet.Cells(1, srcSheet.Columns.Count).End(xlToLeft).Column

    ' データを転写する（範囲を一度にコピー）
    srcSheet.Range(srcSheet.Cells(firstRow, firstCol), srcSheet.Cells(lastRow, lastCol)).Copy _
        Destination:=destSheet.Cells(7, 1)

    ' 取込み元ファイルを閉じる
    srcWorkbook.Close False

    ' 後処理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    MsgBox "データ転写が完了しました。"
End Sub
