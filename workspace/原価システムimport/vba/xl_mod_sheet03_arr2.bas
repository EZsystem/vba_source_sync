Attribute VB_Name = "sheet03_arr2"
Option Explicit

' このコードは、シート "G2_原価S加工データ" の A 列から始まり、6 行の最終列までの範囲を配列 temp1 に取り込み、
' 次に実行するコードに配列を渡すために使用されます。
' 本コードは、本体コード sh03arr_mother02 から呼び出されます。

Sub sh03arr2_mother01()
    
' 配列にデータを取り込む
    Dim temp1() As Variant: Call sh03arrIN_target2(temp1)
'不要列の削除
    Dim temp2() As Variant: Call sh3_arrangementA1(temp1, temp2)
'不要行の削除
    Dim temp3() As Variant: Call sh3_arrangementA2(temp2, temp3)
'セルへの出力
    Call sh3_arroutA(temp3)

    
End Sub



Sub sh03arrIN_target2(ByRef temp1() As Variant)
    Dim ws As Worksheet
    Dim lastCol As Long

    ' シート "G2_原価S加工データ" を設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")
    
    ' 6行の最終列を取得
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    
    ' データ範囲を配列に取り込む
    temp1 = ws.Range(ws.Cells(3, 1), ws.Cells(6, lastCol)).value
End Sub


' このコードは、配列 temp1 を使用して1行目の値が "×" の列を削除し、
' 処理条件と一致するサイズの配列を作成して、その配列に値を入れた後、配列 temp2 として渡します。

Sub sh3_arrangementA1(ByRef temp1() As Variant, ByRef temp2() As Variant)
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long, j As Long, newColCount As Long
    Dim columnsToDelete() As Boolean

    ' 配列の行数と列数を取得
    rowCount = UBound(temp1, 1)
    colCount = UBound(temp1, 2)

    ' 削除する列をマークする配列を初期化
    ReDim columnsToDelete(1 To colCount)
    newColCount = 0

    ' 1行目の値が "×" の列をマーク
    For j = 1 To colCount
        If temp1(1, j) = "×" Then
            columnsToDelete(j) = True
        Else
            columnsToDelete(j) = False
            newColCount = newColCount + 1
        End If
    Next j

    ' 新しい配列 temp2 のサイズを決定
    ReDim temp2(1 To rowCount, 1 To newColCount)

    ' temp1 のデータを temp2 にコピー (削除しない列のみ)
    newColCount = 1
    For j = 1 To colCount
        If Not columnsToDelete(j) Then
            For i = 1 To rowCount
                temp2(i, newColCount) = temp1(i, j)
            Next i
            newColCount = newColCount + 1
        End If
    Next j
End Sub


' このコードは、配列 temp2 を使用して1〜3行目を削除し、
' 新しい配列 temp3 に値を入れて渡します。

Sub sh3_arrangementA2(ByRef temp2() As Variant, ByRef temp3() As Variant)
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long, j As Long
    Dim newRowCount As Long

    ' 配列の行数と列数を取得
    rowCount = UBound(temp2, 1)
    colCount = UBound(temp2, 2)

    ' 新しい配列 temp3 の行数を決定 (1〜3行目を削除)
    newRowCount = rowCount - 3

    ' 新しい配列 temp3 のサイズを決定
    ReDim temp3(1 To newRowCount, 1 To colCount)

    ' temp2 のデータを temp3 にコピー (1〜3行目を除外)
    For i = 4 To rowCount
        For j = 1 To colCount
            temp3(i - 3, j) = temp2(i, j)
        Next j
    Next i
End Sub


' このコードは、既存の配列 temp3 を使用し、シート "G3_原価Sエラー調査" のA列の最終行の1行下に出力します。

Sub sh3_arroutA(ByRef temp3() As Variant)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim targetRange As Range

    ' シート "G3_原価Sエラー調査" を設定
    Set ws = ThisWorkbook.Sheets("G3_原価Sエラー調査")

    ' A列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' 配列の行数と列数を取得
    rowCount = UBound(temp3, 1)
    colCount = UBound(temp3, 2)

    ' 出力範囲を設定
    Set targetRange = ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(lastRow + rowCount, colCount))

    ' 配列を範囲に出力
    targetRange.value = temp3
End Sub
