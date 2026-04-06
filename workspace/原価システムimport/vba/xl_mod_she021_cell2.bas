Attribute VB_Name = "she021_cell2"
Option Explicit


Sub she021cell2_mother01()
'既存データクリア
    'Call sheet021_clear
'基本工事コード記入
    Call she21_cell2_1
'G2参考データコピー
    Call she21_cell2_2
'I22参考データコピー
    Call she21_cell2_3
'予定の削除
    Call she21_cell2_4
End Sub



Sub sheet021_clear()
    '概要: 指定範囲をクリアする自動化スクリプト
    'シート名: G22_原価S基本工事
    '列: A列から始まり、6行目の最終列と同じ
    '行: 7行から始まり、A列の最終行と同じ

    Dim ws As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long
    Dim clearRange As Range
    
    ' シートを設定
    Set ws = ThisWorkbook.Sheets("G22_原価S基本工事")
    
    ' 6行目の最終列を特定
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    
    ' A列の最終行を特定
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' 7行目以降に値があるか確認
    If lastRow >= 7 Then
        ' クリア範囲を設定
        Set clearRange = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, lastCol))
        
        ' 値をクリア
        clearRange.ClearContents
    End If
    
    ' フィルタが設定されている場合はクリア
    If ws.AutoFilterMode Then
        ws.AutoFilter.ShowAllData
    End If
End Sub


Sub she21_cell2_1()
    '概要: コピー元からコピー先に値を貼り付けて、重複を削除する自動化スクリプト
    'コピー元: シート名 G2_原価S加工データ
    'コピー先: シート名 G22_原価S基本工事
    '重複削除: A列の値の重複を削除する

    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim srcLastRow As Long
    Dim destLastRow As Long
    Dim srcRange As Range
    Dim destRange As Range
    Dim destStartRow As Long

    ' シートを設定
    Set wsSrc = ThisWorkbook.Sheets("G2_原価S加工データ")
    Set wsDest = ThisWorkbook.Sheets("G22_原価S基本工事")
    
    ' パフォーマンス最適化の設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' コピー元の最終行を特定
    srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row
    
    ' コピー先の開始行を設定
    destStartRow = 7
    
    ' コピー元の範囲を設定
    Set srcRange = wsSrc.Range(wsSrc.Cells(7, 1), wsSrc.Cells(srcLastRow, 1))
    
    ' コピー先の範囲を設定
    destLastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row
    Set destRange = wsDest.Cells(destStartRow, 1)
    
    ' コピー元の範囲をコピー先に貼り付け
    srcRange.Copy Destination:=destRange
    
    ' コピー先の最終行を再取得
    destLastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row
    
    ' 重複削除の範囲を設定
    Set destRange = wsDest.Range(wsDest.Cells(destStartRow, 1), wsDest.Cells(destLastRow, 1))
    
    ' 重複を削除
    destRange.RemoveDuplicates Columns:=1, header:=xlNo
    
    ' パフォーマンス設定を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub



Sub she21_cell2_2()
    '概要: 条件が一致した行を指定範囲にコピー元からコピー先に値を貼りつける自動化スクリプト
    'コピー元: シート名 G2_原価S加工データ
    'コピー先: シート名 G22_原価S基本工事

    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim srcLastRow As Long
    Dim destLastRow As Long
    Dim srcRow As Long
    Dim destRow As Long
    Dim destStartRow As Long
    Dim srcRange As Range
    Dim destRange As Range
    Dim srcKey As String
    Dim destKey As String

    ' シートを設定
    Set wsSrc = ThisWorkbook.Sheets("G2_原価S加工データ")
    Set wsDest = ThisWorkbook.Sheets("G22_原価S基本工事")
    
    ' パフォーマンス最適化の設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' コピー元の最終行を特定
    srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row
    
    ' コピー先の最終行を特定
    destLastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row
    
    ' コピー先の開始行を設定
    destStartRow = 7
    
    ' コピー先のA列の値をループ
    For destRow = destStartRow To destLastRow
        destKey = wsDest.Cells(destRow, 1).value
        
        ' コピー元のA列の値をループ
        For srcRow = 7 To srcLastRow
            srcKey = wsSrc.Cells(srcRow, 1).value
            
            ' 条件1: コピー先のA列の値とコピー元のA列の値が一致
            ' 条件2: コピー元のB列の値がブランク
            If destKey = srcKey And wsSrc.Cells(srcRow, 2).value = "" Then
                ' コピー元の範囲を設定（B列からX列）
                Set srcRange = wsSrc.Range(wsSrc.Cells(srcRow, 2), wsSrc.Cells(srcRow, 24))
                
                ' コピー先の範囲を設定（B列からX列）
                Set destRange = wsDest.Range(wsDest.Cells(destRow, 2), wsDest.Cells(destRow, 24))
                
                ' コピー元範囲をコピー先に貼り付け
                destRange.value = srcRange.value

            End If
        Next srcRow
    Next destRow
    
    ' パフォーマンス設定を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub



Sub she21_cell2_3()
    '概要: 条件が一致した行を指定範囲にコピー元からコピー先に値を貼りつける自動化スクリプト
    'コピー元: シート名 I22_Icube加工ALL
    'コピー先: シート名 G22_原価S基本工事

    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim srcLastRow As Long
    Dim destLastRow As Long
    Dim srcRow As Long
    Dim destRow As Long
    Dim destStartRow As Long
    Dim srcKey As String
    Dim destKey As String

    ' シートを設定
    Set wsSrc = ThisWorkbook.Sheets("I22_Icube加工ALL")
    Set wsDest = ThisWorkbook.Sheets("G22_原価S基本工事")
    
    ' パフォーマンス最適化の設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' コピー元の最終行を特定
    srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row
    
    ' コピー先の最終行を特定
    destLastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).row
    
    ' コピー先の開始行を設定
    destStartRow = 7
    
    ' コピー先のA列の値をループ
    For destRow = destStartRow To destLastRow
        destKey = wsDest.Cells(destRow, 1).value
        
        ' コピー元のA列の値をループ
        For srcRow = 7 To srcLastRow
            srcKey = wsSrc.Cells(srcRow, 1).value
            
            ' 条件1: コピー先のA列の値とコピー元のA列の値が一致
            If destKey = srcKey Then
                ' コピー元の範囲を設定（AE列）
                wsDest.Cells(destRow, 25).value = wsSrc.Cells(srcRow, 31).value ' コピー先のY列は25列目、コピー元のAE列は31列目
            End If
        Next srcRow
    Next destRow
    
    ' パフォーマンス設定を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub



' このコードは条件が一致したとき行を削除し、削除した件数をメッセージで表示する
Sub she21_cell2_4()
    ' 変数の定義
    Dim wsSurvey As Worksheet
    Dim surveyRow As Long
    Dim surveyLastRow As Long
    Dim deleteCount As Long
    Dim surveyValue As String

    ' 初期設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' シートの設定
    Set wsSurvey = ThisWorkbook.Sheets("G2_原価S加工データ")

    ' 最終行の取得
    surveyLastRow = wsSurvey.Cells(wsSurvey.Rows.Count, 2).End(xlUp).row

    ' 削除した行のカウント初期化
    deleteCount = 0

    ' 調査データの条件を満たす行の削除
    For surveyRow = surveyLastRow To 7 Step -1
        surveyValue = wsSurvey.Cells(surveyRow, 6).value
        If surveyValue = "予定" Then
            wsSurvey.Rows(surveyRow).Delete
            deleteCount = deleteCount + 1
        End If
    Next surveyRow

    ' 後処理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    ' 削除した件数をメッセージで表示
    MsgBox "予定が合ったため該当の　" & deleteCount & " 行を削除ました。"
End Sub
