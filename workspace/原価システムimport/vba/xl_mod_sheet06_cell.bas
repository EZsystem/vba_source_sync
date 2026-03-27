Attribute VB_Name = "sheet06_cell"
Option Explicit

' このコードは条件が一致したとき値を記入し、置き換えた値をメッセージで表示する
Sub sheet06_cell1_1()
    ' 変数の定義
    Dim wsSurvey As Worksheet
    Dim wsComparison As Worksheet
    Dim surveyRow As Long
    Dim surveyLastRow As Long
    Dim comparisonRow As Long
    Dim comparisonLastRow As Long
    Dim surveyValue As String
    Dim comparisonValue As String
    Dim replacedValues As String

    ' 初期設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' シートの設定
    Set wsSurvey = ThisWorkbook.Sheets("G2_原価S加工データ")
    Set wsComparison = ThisWorkbook.Sheets("G6_原価S枝番修正リスト")

    ' 最終行の取得
    surveyLastRow = wsSurvey.Cells(wsSurvey.Rows.Count, 2).End(xlUp).row
    comparisonLastRow = wsComparison.Cells(wsComparison.Rows.Count, 2).End(xlUp).row

    ' 置き換えた値を保持するための変数
    replacedValues = "置き換えた値:" & vbCrLf

    ' 調査データの比較
    For surveyRow = 7 To surveyLastRow
        surveyValue = wsSurvey.Cells(surveyRow, 3).value
        For comparisonRow = 7 To comparisonLastRow
            comparisonValue = wsComparison.Cells(comparisonRow, 2).value
            If surveyValue = comparisonValue Then
                wsSurvey.Cells(surveyRow, 3).value = wsComparison.Cells(comparisonRow, 4).value
                replacedValues = replacedValues & "行 " & surveyRow & ": " & wsComparison.Cells(comparisonRow, 4).value & vbCrLf
                Exit For
            End If
        Next comparisonRow
    Next surveyRow

    ' 後処理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    ' 置き換えた値をメッセージで表示
    MsgBox replacedValues
End Sub
