Attribute VB_Name = "sheet11_arr1"
Option Explicit

Sub she11_arr_mother01()
    '概要: サブコードを呼び出し、データを配列に入れて処理を実行する
    Call sheet11_arr_clear1 '既設ﾃﾞｰﾀクリア
    Dim tmpe1 As Variant: Call sh11arrIN_tmpe1(tmpe1) '配列にﾃﾞｰﾀを入れる
    Dim tmpe2 As Variant: Call she11_arr1(tmpe1, tmpe2) '不要列削除
    Dim tmpe3 As Variant: Call she11_arr2(tmpe2, tmpe3) '不要行削除
    Call sheet11_arr_out1(tmpe3)

End Sub

'既設ﾃﾞｰﾀクリア
Sub sheet11_arr_clear1()
    '概要: 指定範囲をクリアする自動化スクリプト
    'シート名: S1_受注、完工、既払い

    Dim ws As Worksheet
    Dim lastRow As Long

    ' シートを設定
    Set ws = ThisWorkbook.Sheets("S1_受注、完工、既払い")
    
    ' パフォーマンス最適化の設定
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' フィルタをクリア
    If ws.AutoFilterMode Then
        ws.AutoFilter.ShowAllData
    End If
    
    ' AI7セルに値があるか確認
    If ws.Range("AI7").value <> "" Then
        ' AI列の最終行を特定
        lastRow = ws.Cells(ws.Rows.Count, 35).End(xlUp).row
        
        ' 指定範囲をクリア
        ws.Range(ws.Cells(7, 35), ws.Cells(lastRow, 44)).ClearContents
    End If
    
    ' パフォーマンス設定を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub



'配列にﾃﾞｰﾀを入れる
Sub sh11arrIN_tmpe1(tmpe1 As Variant)
    '概要: データを配列に入れ、次に実行するコードに配列を渡す
    'シート名: I22_Icube加工ALL
    '列: A列から始まり、6行の最終列まで
    '行: 6行から始まり、A列の最終行まで

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    ' シートを設定
    Set ws = ThisWorkbook.Sheets("I22_Icube加工ALL")
    
    ' 最終行と最終列を特定
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    
    ' データ範囲を設定
    Set dataRange = ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol))
    
    ' データを配列に格納
    tmpe1 = dataRange.value
End Sub

Sub she11_arr1(tmpe1 As Variant, temp2 As Variant)
    '概要: 既存の配列tmpe1から指定された項目を抽出し、temp2として渡す
    '配列の1行目を基準に処理を行う

    Dim headers As Variant
    Dim sortOrder() As Long
    Dim headerIndex As Long
    Dim i As Long, j As Long
    Dim colCount As Long

    ' 優先項目名の配列
    headers = Array("工事コード", "工事枝番", "追加工事名称", _
                    "工事価格", "粗利益額", "作業所名" & Chr(10) & "建築RN有り", _
                    "受注期", "受注Q", "受注月", "一件工事判定")
    
    ' ソート順序を設定
    ReDim sortOrder(LBound(headers) To UBound(headers))
    For i = LBound(headers) To UBound(headers)
        On Error Resume Next
        headerIndex = Application.Match(headers(i), Application.Index(tmpe1, 1, 0), 0)
        On Error GoTo 0
        If IsNumeric(headerIndex) Then
            sortOrder(i) = headerIndex
        Else
            sortOrder(i) = 0 ' 見つからない場合は0を設定
        End If
    Next i
    
    ' 新しい配列を作成し、指定された項目のみをコピー
    colCount = 1
    ReDim temp2(LBound(tmpe1, 1) To UBound(tmpe1, 1), 1 To UBound(headers) + 1) ' 列範囲を1から11に設定
    For i = LBound(sortOrder) To UBound(sortOrder)
        If sortOrder(i) > 0 Then
            For j = LBound(tmpe1, 1) To UBound(tmpe1, 1)
                temp2(j, colCount) = tmpe1(j, sortOrder(i))
            Next j
            colCount = colCount + 1
        End If
    Next i
End Sub

Sub she11_arr2(tmpe2 As Variant, temp3 As Variant)
    '概要: 既存の配列tmpe2から指定された行を削除し、temp3として渡す
    '削除条件に基づいて新しい配列を作成する

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tempArray() As Variant
    Dim i As Long, j As Long, k As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim deleteCount As Long

    ' シートを設定
    Set ws = ThisWorkbook.Sheets("S1_受注、完工、既払い")

    ' 配列のサイズを取得
    rowCount = UBound(tmpe2, 1)
    colCount = UBound(tmpe2, 2)

    ' 削除する行数をカウント
    deleteCount = 0
    For i = 1 To rowCount
        ' 処理条件1: 1列目の行を削除
        If tmpe2(i, 1) = "" Then
            deleteCount = deleteCount + 1
        ' 処理条件2: 7列目の値がD1の値と異なる行を削除
        ElseIf tmpe2(i, 7) <> ws.Range("D1").value Then
            deleteCount = deleteCount + 1
        ' 処理条件3: 10列目の値が"小口工事"の行を削除
        ElseIf tmpe2(i, 10) = "小口工事" Then
            deleteCount = deleteCount + 1
        ' 新しい処理条件: 6列目の値が"建築部RN"の行を削除
        ElseIf tmpe2(i, 6) = "建築部RN" Then
            deleteCount = deleteCount + 1
        End If
    Next i

    ' 新しい配列のサイズを決定
    ReDim tempArray(1 To rowCount - deleteCount, 1 To colCount)

    ' 新しい配列に値をコピー
    j = 1
    For i = 1 To rowCount
        ' 削除条件に該当しない行を新しい配列にコピー
        If Not (tmpe2(i, 1) = "" Or tmpe2(i, 7) <> ws.Range("D1").value Or tmpe2(i, 10) = "小口工事" Or tmpe2(i, 6) = "建築部RN") Then
            For k = 1 To colCount
                tempArray(j, k) = tmpe2(i, k)
            Next k
            j = j + 1
        End If
    Next i

    ' 新しい配列をtemp3として渡す
    temp3 = tempArray
End Sub


Sub sheet11_arr_out1(tmpe3 As Variant)
    '概要: 既存の配列tmpe3を指定する位置に出力する
    'シート名: S1_受注、完工、既払い

    Dim ws As Worksheet
    Dim startCell As Range
    Dim endCell As Range
    Dim lastRow As Long
    Dim lastCol As Long

    ' シートを設定
    Set ws = ThisWorkbook.Sheets("S1_受注、完工、既払い")
    
    ' 出力開始位置を設定
    Set startCell = ws.Range("AI7")
    
    ' 配列のサイズを取得
    lastRow = UBound(tmpe3, 1)
    lastCol = UBound(tmpe3, 2)
    
    ' 出力範囲を設定
    Set endCell = startCell.Offset(lastRow - 1, lastCol - 1)
    
    ' 配列をシートに出力
    ws.Range(startCell, endCell).value = tmpe3
End Sub
