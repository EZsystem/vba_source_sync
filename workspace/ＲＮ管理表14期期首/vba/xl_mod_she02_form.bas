Attribute VB_Name = "she02_form"
Option Explicit

'予実績表のセル塗りつぶし
Sub she02_form_All()
    ' 画面更新、計算設定、警告表示を無効化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False


'セルの書式(セル塗りつぶし)コピー(予実績管理表)
    Call she02_formCopy1
'集計列塗りつぶすための対象外行非表示
    Call she02_form_rowInvisible1
'集計列のセル塗りつぶしコピー(予実績管理表)
    Call she02_formCopy2
'行の非常時解除
    Call she02_form_rowInvisible2
    
    ' 画面更新、計算設定、警告表示を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True


MsgBox "予実績表のセル塗りつぶし完了しました。", vbInformation
End Sub



'行の非常時解除
Sub she02_form_rowInvisible2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("各県RNQ毎")

    ' 最終行を取得（任意の列を基準にする、ここではA列）
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' 行の非表示設定を解除
    For i = 1 To lastRow
        If ws.Rows(i).Hidden = True Then
            ws.Rows(i).Hidden = False ' 非表示を解除
        End If
    Next i

    'MsgBox "行の非表示設定を解除しました。", vbInformation
End Sub



'集計列のセル塗りつぶしコピー(予実績管理表)
Sub she02_formCopy2()
    Dim ws As Worksheet
    Dim copyStartCols As Collection
    Dim pasteStartRow As Long, pasteEndRow As Long
    Dim targetRow As Long
    Dim lastCol As Long, lastRow As Long
    Dim i As Long, col As Variant

    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("各県RNQ毎")

    ' コピー元の列を特定（"集計塗り"がある列をすべて収集）
    Set copyStartCols = New Collection
    targetRow = 18 ' コピー元の行（18行目）
    lastCol = ws.Cells(targetRow, ws.Columns.Count).End(xlToLeft).Column ' 最終列を取得

    For i = 53 To lastCol ' BA列(53列目)から最終列まで
        If ws.Cells(targetRow, i).value = "集計塗り" Then
            copyStartCols.Add i ' 集計塗りの列番号を記録
        End If
    Next i

    ' コピー先の行範囲を特定（"集計書式コピー始り"から"集計書式コピー終わり"）
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row ' B列の最終行を取得
    For i = 30 To lastRow
        If ws.Cells(i, 2).value = "集計書式コピー始り" Then
            pasteStartRow = i
        ElseIf ws.Cells(i, 2).value = "集計書式コピー終わり" Then
            pasteEndRow = i
            Exit For
        End If
    Next i

    ' 範囲が正しく設定されているかチェック
    If copyStartCols.Count = 0 Or pasteStartRow = 0 Or pasteEndRow = 0 Then
        MsgBox "コピー範囲または貼り付け範囲が正しく設定されていません。", vbExclamation
        Exit Sub
    End If

    ' 色をコピー（非表示の行をスキップし、列ごとに処理）
    Dim copyRange As Range, pasteRange As Range
    For Each col In copyStartCols
        For i = pasteStartRow To pasteEndRow
            ' 非表示行はスキップ
            If ws.Rows(i).Hidden = False Then
                ' コピー元のセル
                Set copyRange = ws.Cells(targetRow, col)
                ' コピー先のセル
                Set pasteRange = ws.Cells(i, col)

                ' 色をコピー
                pasteRange.Interior.Color = copyRange.Interior.Color
            End If
        Next i
    Next col

    ' コピー範囲の選択を解除
    ws.Cells(1, 1).Select

    'MsgBox "書式コピーが完了しました（非表示行を除く）。", vbInformation
End Sub




'集計列塗りつぶすための対象外行非表示
Sub she02_form_rowInvisible1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim checkCol As Long
    Dim startRow As Long
    Dim i As Long

    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("各県RNQ毎")

    ' 調査範囲を設定
    checkCol = 2 ' B列
    startRow = 30 ' 調査開始行
    lastRow = ws.Cells(ws.Rows.Count, checkCol).End(xlUp).row ' B列の最終行を取得

    ' 指定条件に基づいて行を非表示にする
    For i = startRow To lastRow
        Select Case ws.Cells(i, checkCol).value
            Case "書式コピー始り", "集計塗り対象外", "書式コピー終わり"
                ws.Rows(i).Hidden = True ' 行を非表示に設定
        End Select
    Next i

    'MsgBox "指定条件に基づいて行を非表示にしました。", vbInformation
End Sub



'セルの書式(セル塗りつぶし)コピー(予実績管理表)
Sub she02_formCopy1()
    Dim ws As Worksheet
    Dim copyStartRow As Long, copyEndRow As Long
    Dim pasteStartCol As Long, pasteEndCol As Long
    Dim copyCol As Long, startRow As Long
    Dim i As Long, j As Long

    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("各県RNQ毎")

    ' コピー元の列と条件を設定
    copyCol = 2 ' B列
    startRow = 20 ' 調べる開始行
    For i = startRow To ws.Cells(ws.Rows.Count, copyCol).End(xlUp).row
        If ws.Cells(i, copyCol).value = "書式コピー始り" Then
            copyStartRow = i
        ElseIf ws.Cells(i, copyCol).value = "書式コピー終わり" Then
            copyEndRow = i
            Exit For
        End If
    Next i

    ' コピー先の列方向を調べる
    Dim searchRow As Long
    searchRow = 18 ' AH列を含む行

    For j = 34 To ws.Columns.Count ' AH列(34列目)から右方向に調べる
        If ws.Cells(searchRow, j).value = "書式コピー始り1" Then
            pasteStartCol = j
        ElseIf ws.Cells(searchRow, j).value = "書式コピー終わり1" Then
            pasteEndCol = j
            Exit For
        End If
    Next j

    ' 範囲が正しく設定されているかチェック
    If copyStartRow = 0 Or copyEndRow = 0 Or pasteStartCol = 0 Or pasteEndCol = 0 Then
        MsgBox "書式コピーの範囲が正しく設定されていません。", vbExclamation
        Exit Sub
    End If

    ' 書式をコピーする
    Dim copyRange As Range, pasteRange As Range
    Dim pasteRow As Long
    For i = copyStartRow To copyEndRow
        For j = pasteStartCol To pasteEndCol
            ' コピー元のセル
            Set copyRange = ws.Cells(i, copyCol)

            ' コピー先のセル
            Set pasteRange = ws.Cells(i, j)

            ' 書式のコピー
            With pasteRange
                .Interior.Color = copyRange.Interior.Color ' 背景色
                '.Borders(xlEdgeLeft).LineStyle = copyRange.Borders(xlEdgeLeft).LineStyle ' 左罫線
                '.Borders(xlEdgeTop).LineStyle = copyRange.Borders(xlEdgeTop).LineStyle ' 上罫線
                '.Borders(xlEdgeBottom).LineStyle = copyRange.Borders(xlEdgeBottom).LineStyle ' 下罫線
                '.Borders(xlEdgeRight).LineStyle = copyRange.Borders(xlEdgeRight).LineStyle ' 右罫線
            End With
        Next j
    Next i

    ' コピー範囲の選択を解除
    ws.Cells(1, 1).Select

    'MsgBox "書式コピーが完了しました。", vbInformation
End Sub

