Attribute VB_Name = "sheet03_arr"
Option Explicit

' 本コードはシート「G2_原価S加工データ」のデータを配列に取り込み、
' 配列名「tmpe1」として次のコードに渡すためのものです。
' このコードは本体コード「sh03arr_mother01」から呼び出されます。

Sub sh03arr_mother01()
'配列に値を入れる
    Dim tmpe1() As Variant: Call sh03arrIN_target(tmpe1)
'配列の不要列削除
    Dim tmpe2() As Variant: Call sh3_arrangement01(tmpe1, tmpe2)
'配列の不要行削除1：上部
    Dim tmpe3() As Variant: Call sh3_arrangement02(tmpe2, tmpe3)
'配列の不要行削除2：エラーが無い行を削除
    Dim tmpe4() As Variant: Call sh3_arrangement03(tmpe3, tmpe4)
'除外工事コードの行を削除
    Dim tmpe5() As Variant: Call sh3_arrangement04(tmpe4, tmpe5)
'セルへの出力
    Call sh3_arrout(tmpe5)

End Sub

Sub sh03arrIN_target(ByRef tmpe1() As Variant)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long

    ' 対象シートを設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")

    ' 最終行と最終列を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column

    ' データ範囲を配列に取り込み
    tmpe1 = ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol)).value
End Sub


' このコードは、既存の配列 temp1 を使用し、1行目の値が "×" の列を削除し、
' その結果を新しい配列 temp2 に渡します。

Sub sh3_arrangement01(ByRef temp1() As Variant, ByRef temp2() As Variant)
    Dim rowCount As Long
    Dim colCount As Long
    Dim newColCount As Long
    Dim i As Long, j As Long, k As Long
    Dim colsToDelete() As Boolean
    
    ' 配列の行数と列数を取得
    rowCount = UBound(temp1, 1)
    colCount = UBound(temp1, 2)
    
    ' 削除対象の列をマークするためのブール配列を初期化
    ReDim colsToDelete(1 To colCount)
    newColCount = 0
    
    ' 1行目をチェックして "×" の列をマーク
    For j = 1 To colCount
        If temp1(1, j) = "×" Then
            colsToDelete(j) = True
        Else
            colsToDelete(j) = False
            newColCount = newColCount + 1
        End If
    Next j
    
    ' 新しい配列 temp2 のサイズを決定
    ReDim temp2(1 To rowCount, 1 To newColCount)
    
    ' temp1 のデータを temp2 にコピー（"×" の列を除く）
    k = 1
    For j = 1 To colCount
        If Not colsToDelete(j) Then
            For i = 1 To rowCount
                temp2(i, k) = temp1(i, j)
            Next i
            k = k + 1
        End If
    Next j
End Sub


' このコードは、既存の配列 temp2 を使用し、1から4行目を削除し、
' その結果を新しい配列 temp3 に渡します。

Sub sh3_arrangement02(ByRef temp2() As Variant, ByRef temp3() As Variant)
    Dim rowCount As Long
    Dim colCount As Long
    Dim newRowCount As Long
    Dim i As Long, j As Long

    ' 配列の行数と列数を取得
    rowCount = UBound(temp2, 1)
    colCount = UBound(temp2, 2)
    
    ' 新しい行数を計算（1から4行目を削除するため）
    newRowCount = rowCount - 4

    ' 新しい配列 temp3 のサイズを決定
    ReDim temp3(1 To newRowCount, 1 To colCount)

    ' temp2 のデータを temp3 にコピー（1から4行目を除く）
    For i = 1 To newRowCount
        For j = 1 To colCount
            temp3(i, j) = temp2(i + 4, j)
        Next j
    Next i
End Sub





' このコードは、既存の配列 temp5 を使用し、シート "G3_原価Sエラー調査" の A 列の最終行の1段下に出力します。
Sub sh3_arrout(ByRef temp4() As Variant)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputRange As Range
    
    ' temp4 に値が入っていない場合は処理を中断
    If (Not temp4) = -1 Then
        MsgBox "temp4に値が入っていません。処理を中断します。"
        Exit Sub
    End If

    ' シート "G3_原価Sエラー調査" を設定
    Set ws = ThisWorkbook.Sheets("G3_原価Sエラー調査")
    
    ' A列の最終行の1段下を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    
    ' 出力範囲を設定
    Set outputRange = ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow + UBound(temp4, 1) - 1, UBound(temp4, 2)))
    
    ' 配列のデータを範囲に出力
    outputRange.value = temp4
End Sub

' このコードは、既存の配列 temp3 を使用し、15列が空文字列かつ18〜22列の値がブランクの場合に行を削除し、
' その結果を新しい配列 temp4 に渡します。
Sub sh3_arrangement03(ByRef temp3() As Variant, ByRef temp4() As Variant)
    Dim rowCount As Long
    Dim colCount As Long
    Dim newRowCount As Long
    Dim i As Long, j As Long
    Dim currentRow As Long
    Dim validRowCount As Long
    
    ' 配列の行数と列数を取得
    rowCount = UBound(temp3, 1)
    colCount = UBound(temp3, 2)
    
    ' 有効な行数をカウント（15列が空文字列でない、かつ18〜22列がブランクでない行）
    validRowCount = 0
    For i = 1 To rowCount
        If Not (temp3(i, 15) = "" And IsEmpty(temp3(i, 18)) And IsEmpty(temp3(i, 19)) And _
                IsEmpty(temp3(i, 20)) And IsEmpty(temp3(i, 21)) And IsEmpty(temp3(i, 22))) Then
            validRowCount = validRowCount + 1
        End If
    Next i
    
    ' 新しい配列 temp4 のサイズを決定
    ReDim temp4(1 To validRowCount, 1 To colCount)
    
    ' temp3 のデータを temp4 にコピー（15列が空文字列でない、かつ18〜22列がブランクでない行のみ）
    currentRow = 1
    For i = 1 To rowCount
        If Not (temp3(i, 15) = "" And IsEmpty(temp3(i, 18)) And IsEmpty(temp3(i, 19)) And _
                IsEmpty(temp3(i, 20)) And IsEmpty(temp3(i, 21)) And IsEmpty(temp3(i, 22))) Then
            For j = 1 To colCount
                temp4(currentRow, j) = temp3(i, j)
            Next j
            currentRow = currentRow + 1
        End If
    Next i
End Sub



' このコードは、配列 temp4 の行を削除し、新しい配列 temp5 として渡す自動化スクリプトを実行します
Sub sh3_arrangement04(temp4 As Variant, ByRef temp5() As Variant)
    ' 変数の定義
    Dim ws As Worksheet
    Dim temp5Index As Long
    Dim temp4Row As Long
    Dim temp4Col As Long
    Dim i As Long
    Dim j As Long
    Dim exclusionValues As Variant
    Dim matchFound As Boolean
    Dim countValidRows As Long
    Dim lastRow As Long

    ' シートの設定
    Set ws = ThisWorkbook.Sheets("G7_エラー値調査除外工事")

    ' 除外する値を取得
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).row
    exclusionValues = ws.Range(ws.Cells(7, 3), ws.Cells(lastRow, 3)).value

    ' temp4 の行数と列数を取得
    temp4Row = UBound(temp4, 1)
    temp4Col = UBound(temp4, 2)

    ' 条件に一致しない行数をカウント
    countValidRows = 0
    For i = 1 To temp4Row
        matchFound = False
        For j = 1 To UBound(exclusionValues, 1)
            ' temp4の3列目とexclusionValuesをトリムして比較
            If Trim(CStr(temp4(i, 3))) = Trim(CStr(exclusionValues(j, 1))) Then
                matchFound = True
                Exit For
            End If
        Next j
        If Not matchFound Then
            countValidRows = countValidRows + 1
        End If
    Next i

    ' 条件に一致しない行がない場合のメッセージ表示
    If countValidRows = 0 Then
        MsgBox "除外工事コードはありませんでした"
        Exit Sub
    End If

    ' temp5 配列のサイズを決定
    ReDim temp5(1 To countValidRows, 1 To temp4Col)

    ' temp5 に値を入れる
    temp5Index = 0
    For i = 1 To temp4Row
        matchFound = False
        For j = 1 To UBound(exclusionValues, 1)
            ' temp4の3列目とexclusionValuesをトリムして比較
            If Trim(CStr(temp4(i, 3))) = Trim(CStr(exclusionValues(j, 1))) Then
                matchFound = True
                Exit For
            End If
        Next j
        If Not matchFound Then
            temp5Index = temp5Index + 1
            For j = 1 To temp4Col
                temp5(temp5Index, j) = temp4(i, j)
            Next j
        End If
    Next i
End Sub
