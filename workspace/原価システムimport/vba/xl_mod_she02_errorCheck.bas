Attribute VB_Name = "she02_errorCheck"
Option Explicit

Sub sh2_errochek() 'シート：G2_原価S加工データのエラー関係記入
    Call sh02_errorCheck01 '避難関数のコピーと値化
    Call sh02_errorCheck1 '原価Sから重複取込みチェック
    Call sh02_errorCheck2 '受注当初粗利率記入
    Call sh02_errorCheck3 '現状粗利上限エラーチェック
    Call sh02_errorCheck4 '現状粗利下限エラーチェック
    Call sh02_errorCheck45 '現状粗利上下限エラーチェック
    Call sh02_errorCheck5 '完工後の一定期日経過後の支払有無チェック
End Sub


Sub sh02_errorCheck01() '避難関数のコピーと値化
    ' このコードは、シート "G2_原価S加工データ" の指定範囲に計算式を貼り付け、値化します。
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim formulaRange As Range
    Dim pasteRange As Range
    
    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 7行目以降にデータが無い場合は終了
    If lastRow < 7 Then
        'MsgBox "7行目以降にデータがありません。", vbInformation
        Exit Sub
    End If
    
    ' コピー元の開始列を取得（"避難関数→"が有る列より1列右隣）
    startCol = ws.Rows(1).Find(What:="避難関数→", LookIn:=xlValues, LookAt:=xlWhole).Column + 1
    
    ' コピー元の終了列を取得（ブランクになるまで）
    endCol = startCol
    Do Until IsEmpty(ws.Cells(1, endCol))
        endCol = endCol + 1
    Loop
    endCol = endCol - 1
    
    ' コピー元範囲を設定
    Set formulaRange = ws.Range(ws.Cells(1, startCol), ws.Cells(1, endCol))
    
    ' コピー先範囲を設定
    Set pasteRange = ws.Range(ws.Cells(7, startCol), ws.Cells(lastRow, endCol))
    
    ' 画面更新を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 計算式を貼り付け
    formulaRange.Copy
    pasteRange.PasteSpecial Paste:=xlPasteFormulas
    
    ' 計算を自動に戻す
    Application.Calculation = xlCalculationAutomatic
    
    ' 貼り付けた計算式を値に変換
    pasteRange.Copy
    pasteRange.PasteSpecial Paste:=xlPasteValues
    
    ' クリップボードをクリア
    Application.CutCopyMode = False
    
    ' 画面更新を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' 完了メッセージ
    'MsgBox "計算式を貼り付け、値に変換しました。", vbInformation
End Sub



' このコードは、シート「G2_原価S加工データ」のC列に重複があるかどうかを調べ、
' 重複があった場合は対応する行のAR列に「重複有り」と記入します。
'原価Sから重複取込みチェック
Sub sh02_errorCheck1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim foundDuplicate As Boolean
    Dim dict As Object

    ' 対象シートを設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Dictionaryオブジェクトを作成
    Set dict = CreateObject("Scripting.Dictionary")

    ' 重複チェック
    For i = 7 To lastRow
        If Not IsEmpty(ws.Cells(i, "C").value) Then
            If dict.Exists(ws.Cells(i, "C").value) Then
                ' 重複があった場合、対応する行のBD列に「重複有り」と記入
                ws.Cells(i, "BD").value = "重複有り"
            Else
                dict.Add ws.Cells(i, "C").value, Nothing
            End If
        End If
    Next i

    ' パフォーマンス最適化のための設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub



Sub sh02_errorCheck2()
    ' 概要: 調査データと参考データを比較し、一致した場合に調査データのBE列に値を記入
    ' 生成日: 2024/11/05
    Dim wsSurvey As Worksheet
    Dim wsReference As Worksheet
    Dim surveyLastRow As Long
    Dim referenceLastRow As Long
    Dim i As Long, j As Long

    ' 調査データと参考データのシートを設定
    Set wsSurvey = ThisWorkbook.Sheets("G2_原価S加工データ")
    Set wsReference = ThisWorkbook.Sheets("I22_Icube加工ALL")

    ' 調査データと参考データの最終行を取得
    surveyLastRow = wsSurvey.Cells(wsSurvey.Rows.Count, "A").End(xlUp).row
    referenceLastRow = wsReference.Cells(wsReference.Rows.Count, "A").End(xlUp).row

    ' 画面更新を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' 調査データと参考データを比較
    For i = 7 To surveyLastRow
        For j = 7 To referenceLastRow
            If wsSurvey.Cells(i, "C").value = wsReference.Cells(j, "C").value Then
                ' 値がブランクの場合をチェックして処理をスキップ
                If IsEmpty(wsReference.Cells(j, "I").value) Or IsEmpty(wsReference.Cells(j, "H").value) Then
                    GoTo nextRow
                End If
                wsSurvey.Cells(i, "BE").value = wsReference.Cells(j, "I").value / wsReference.Cells(j, "H").value
                Exit For
            End If
        Next j
nextRow:
    Next i

    ' 画面更新と計算を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    ' 完了メッセージ
    'MsgBox "比較と記入が完了しました。", vbInformation
End Sub


Sub sh02_errorCheck3()
    '現状粗利上限エラーチェック
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim addValue As Double
    Dim checkValue As Double

    ' 実行するシートの設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' 加算値を取得
    addValue = ws.Range("BF2").value

    ' 画面更新と計算を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' 調査データのK列と参考データのAS列を比較して処理を行う
    For i = 7 To lastRow
        ' 条件1: E列の値が「鈴木　雄太」の場合は処理を行わない
        If ws.Cells(i, "E").value = "鈴木　雄太" Then
            GoTo NextIteration
        End If
        
        ' 条件2: AS列がブランクの場合は処理を行わない
        If IsEmpty(ws.Cells(i, "BE").value) Then
            GoTo NextIteration
        End If
        
        ' AJ列がマイナスの場合、処理をスキップ
        If ws.Cells(i, 11).value < 0 Then GoTo NextIteration
        

        
        ' 比較処理
        checkValue = ws.Cells(i, "BE").value + addValue
        If checkValue < ws.Cells(i, "K").value Then
            ws.Cells(i, "BF").value = "エラー"
        End If
        
NextIteration:
    Next i

    ' 画面更新と計算を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub




' 調査データを参考データと比較し、条件が一致したときに値を記入するコード
'現状粗利下限エラーチェック
Sub sh02_errorCheck4()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim addValue As Double
    Dim kValue As Double
    Dim asValue As Double
    
    ' 画面更新を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' シートを設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 加算値を取得
    addValue = ws.Range("BH2").value
    
    ' 調査データを参考データと比較して値を記入
    For i = 7 To lastRow
        ' E列が"鈴木 雄太"の場合、処理をスキップ
        If ws.Cells(i, "E").value = "鈴木　雄太" Then GoTo NextIteration
        
        ' AS列がブランクの場合、処理をスキップ
        If IsEmpty(ws.Cells(i, "BE").value) Then GoTo NextIteration
        
        ' AJ列がマイナスの場合、処理をスキップ
        If ws.Cells(i, "K").value < 0 Then GoTo NextIteration
        
        kValue = ws.Cells(i, "K").value
        asValue = ws.Cells(i, "BH").value
        
        ' 調査データから加算値を引いた値が参考データより小さいか調べる
        If kValue < (asValue - addValue) Then
            ws.Cells(i, "BH").value = "エラー"
        End If
        
NextIteration:
    Next i
    
    ' 画面更新と計算を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    ' 完了メッセージ
    'MsgBox "エラーチェックが完了しました。", vbInformation
End Sub



'現状粗利上下限エラーチェック
' sh02_errorCheck45: 調査データと参考データを比較し、条件が一致した場合に値を記入するスクリプト
Sub sh02_errorCheck45()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' 画面更新を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' 調査データのシートを設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' データをループして条件をチェック
    For i = 7 To lastRow
        If ws.Cells(i, "BF").value = "エラー" Or ws.Cells(i, "BG").value = "エラー" Then
            ws.Cells(i, "BH").value = "エラー"
        End If
    Next i
    
    ' 画面更新と計算を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    ' 完了メッセージ
    'MsgBox "チェックが完了しました。", vbInformation
End Sub

'完工後の一定期日経過後の支払有無チェック
' sh02_errorCheck5: 調査データの日付に加算値を加えて比較し、条件が一致した場合に値を記入するコード
Sub sh02_errorCheck5()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim surveyDate As Date
    Dim referenceDate As String
    Dim addMonths As Integer
    Dim checkDate As Date

    ' シート名を設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 加算値を取得
    addMonths = ws.Range("BI2").value

    ' 画面更新と計算を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' データをループ処理
    For i = 7 To lastRow
        ' BE列がブランクの場合は処理しない
        If Len(ws.Cells(i, "BE").value) > 0 Then
            surveyDate = ws.Cells(i, "V").value
            referenceDate = ws.Cells(i, "Y").value
            
            ' 調査データの日付に加算値（ヶ月）を足した日付を計算
            checkDate = DateAdd("m", addMonths, surveyDate)
            
            ' 本日より過去の日付であり、かつ参考データに値が記入されている場合
            If checkDate < Date And Len(referenceDate) > 0 Then
                ws.Cells(i, "BI").value = addMonths & "ヶ月経過"
            End If
        End If
    Next i

    ' 画面更新と計算を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub
