Attribute VB_Name = "she02_format"
Option Explicit

Sub FormatCellsBasedOnFourthRow()
    ' このコードは、指定されたシートで、4行目のセルの内容に基づいてセルの書式を設定します。
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    
    ' 実行するシートを設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")
    
    ' 最終行と最終列を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    
    ' 画面更新と計算を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' 各列ごとに4行目のセルの内容に基づいて書式設定を適用
    For i = 1 To lastCol
        Select Case ws.Cells(4, i).value
            Case "書式：数値カンマ"
                On Error Resume Next
                ws.Range(ws.Cells(7, i), ws.Cells(lastRow, i)).NumberFormat = "#,##0;[red]-#,##0"
                On Error GoTo 0
            Case "書式：%"
                On Error Resume Next
                ws.Range(ws.Cells(7, i), ws.Cells(lastRow, i)).NumberFormat = "0.00%"
                On Error GoTo 0
            Case "書式：日付"
                On Error Resume Next
                ws.Range(ws.Cells(7, i), ws.Cells(lastRow, i)).NumberFormat = "yyyy/mm/dd"
                On Error GoTo 0
        End Select
    Next i
    
    ' 画面更新と計算を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub
