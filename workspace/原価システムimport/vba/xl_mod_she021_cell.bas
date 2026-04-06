Attribute VB_Name = "she021_cell"
Option Explicit
'工事名の整形と、県名の抽出
Sub she021_cellMother01()
    Call she021_cell_1 ' 目的: 左から10文字目がブランクの場合、左側をクリア
    Call she021_cell_2 ' 目的: コピー元範囲からコピー先にデータを貼り付け、指定条件に一致する文字列を置換
    Call she021_cell_3 ' 目的: コピー元範囲からコピー先にデータを貼り付け、指定条件に一致する文字列を置換
End Sub

' === she021_cell_3 ===
' 条件に一致する行の一部をコピーし、文字列置換を行う
' 目的: コピー元範囲からコピー先にデータを貼り付け、指定条件に一致する文字列を置換
' 作成日: 2024/11/11

Sub she021_cell_3()
    ' 定義
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String

    ' シート「G22_原価S基本工事」を設定
    Set ws = ThisWorkbook.Sheets("G22_原価S基本工事")

    ' 最終行を取得（D列の最終行）
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).row

    ' === 最適化設定 ===
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' コピー元範囲（D列）からコピー先範囲（F列）に値を貼り付け
    ws.Range("F7:F" & lastRow).value = ws.Range("D7:D" & lastRow).value

    ' コピー先範囲（F列）の値をチェックして置換処理
    For i = 7 To lastRow
        cellValue = ws.Cells(i, "F").value
        
        ' 条件：10文字目が「Ｑ」（全角）の行が対象
        If Mid(cellValue, 10, 1) = "Ｑ" Then
            ' 9,10文字だけを残す
            ws.Cells(i, "F").value = Mid(cellValue, 9, 2)
        End If
    Next i

    ' === 最適化設定を元に戻す ===
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    ' 完了メッセージ
    'MsgBox "データのコピーと置換処理が完了しました。", vbInformation
End Sub

' === she021_cell_2 ===
' 条件に一致する行の一部をコピーし、文字列置換を行う
' 目的: コピー元範囲からコピー先にデータを貼り付け、指定条件に一致する文字列を置換
' 作成日: 2024/11/11

Sub she021_cell_2()
    ' 定義
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String

    ' シート「G22_原価S基本工事」を設定
    Set ws = ThisWorkbook.Sheets("G22_原価S基本工事")

    ' 最終行を取得（D列の最終行）
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).row

    ' === 最適化設定 ===
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' コピー元範囲（D列）からコピー先範囲（E列）に値を貼り付け
    ws.Range("E7:E" & lastRow).value = ws.Range("D7:D" & lastRow).value

    ' コピー先範囲（E列）の値をチェックして置換処理
    For i = 7 To lastRow
        cellValue = ws.Cells(i, "E").value
        
        ' 条件：3〜4文字目が「ＲＮ」（全角）の行が対象
        If Mid(cellValue, 3, 2) = "ＲＮ" Then
            ' 3文字目以降を削除
            ws.Cells(i, "E").value = Left(cellValue, 2)
        End If
    Next i

    ' === 最適化設定を元に戻す ===
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    ' 完了メッセージ
    'MsgBox "データのコピーと置換処理が完了しました。", vbInformation
End Sub



' === she021_cell_1 ===
' 指定セルの値を置換えて出力する
' 目的: 左から10文字目がブランクの場合、左側をクリア
' 日付: 2024/11/08

Sub she021_cell_1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim cellValue As String

    ' 最適化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' 対象シートを設定
    Set ws = ThisWorkbook.Sheets("G22_原価S基本工事")

    ' D列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).row

    ' 対象範囲をループ処理
    For Each cell In ws.Range("D7:D" & lastRow)
        cellValue = cell.value

        ' 条件1：左から10文字目がブランクの場合、左側をクリア
        If Len(cellValue) >= 10 And Mid(cellValue, 10, 1) = " " Then
            cell.value = Mid(cellValue, 11) ' 左側の10文字を削除
        End If
    Next cell

    ' 最適化設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    ' 完了メッセージ
    'MsgBox "指定範囲の値の置換が完了しました。", vbInformation
End Sub
