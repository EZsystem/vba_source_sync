Attribute VB_Name = "sheet21_cell"
Option Explicit

Sub sheet21_cell1()
    '概要: 計算式をコピーし、値に変換する自動化スクリプト
    'シート名: S1_受注、完工、既払い

    Dim ws As Worksheet
    Dim lastCol As Long
    Dim copyStartCol As Long
    Dim copyEndCol As Long
    Dim pasteStartRow As Long
    Dim pasteEndRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim col As Long
    Dim targetCol As Long

    ' シートを設定
    Set ws = ThisWorkbook.Sheets("S1_受注、完工、既払い")
    
    ' パフォーマンス最適化の設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = False
    
    ' 1行目の最終列を特定
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' コピー先の開始行を設定
    pasteStartRow = 7
    
    ' 1行目のA列から最終列までをループ
    startCol = 1
    endCol = lastCol
    For col = startCol To endCol
        If ws.Cells(1, col).value = "関数避難場所→" Then
            copyStartCol = col + 1
            
            ' コピー範囲の最終列を特定
            copyEndCol = copyStartCol
            Do While ws.Cells(1, copyEndCol).value <> "" And copyEndCol <= lastCol
                copyEndCol = copyEndCol + 1
            Loop
            copyEndCol = copyEndCol - 1
            
            ' コピー先の終了行を特定
            If ws.Cells(pasteStartRow, col).value <> "" Then
                pasteEndRow = ws.Cells(ws.Rows.Count, col).End(xlUp).row
            Else
                pasteEndRow = pasteStartRow
            End If
            
            ' 計算式をコピーし、スペシャルペーストで貼り付け
            ws.Range(ws.Cells(1, copyStartCol), ws.Cells(1, copyEndCol)).Copy
            ws.Range(ws.Cells(pasteStartRow, copyStartCol), ws.Cells(pasteEndRow, copyEndCol)).PasteSpecial Paste:=xlPasteFormulas

            ' 貼り付けた範囲を値に変換
            ws.Range(ws.Cells(pasteStartRow, copyStartCol), ws.Cells(pasteEndRow, copyEndCol)).value = ws.Range(ws.Cells(pasteStartRow, copyStartCol), ws.Cells(pasteEndRow, copyEndCol)).value
        End If
    Next col
    
    ' ScreenUpdatingとDisplayAlertsを元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' クリップボードをクリア
    Application.CutCopyMode = False
End Sub
