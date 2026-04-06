Attribute VB_Name = "she02_clearFilter"
Option Explicit

Sub sheet02_clear()
    ' このコードは、シート "G2_原価S加工データ" の指定範囲をクリアし、
    ' フィルタが設定されている場合は解除します。
    ' 7行目以降に値が無い場合は実施しません。

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")
    
    ' 最終行と最終列を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    
    ' 値が無い場合は終了
    If lastRow < 7 Then
        'MsgBox "7行目以降に値がありません。", vbInformation
        Exit Sub
    End If
    
    ' 画面更新と計算を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' フィルタを解除
    If ws.AutoFilterMode Or ws.FilterMode Then
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0
    End If
    
    ' 指定範囲の値をクリア
    ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, lastCol)).ClearContents
    
    ' 画面更新と計算を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    ' 完了メッセージ
    'MsgBox "指定範囲の値をクリアしました。", vbInformation
End Sub


Sub sheet02_filtering()
    ' このコードは、シート "G2_原価S加工データ" にフィルターを設定します。
    ' フィルターが既に設定されている場合でも、再設定を行います。
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' 対象のシートを設定
    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 最終列を取得
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    
    ' 画面更新を無効にしてパフォーマンスを最適化
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' フィルターを解除
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ' フィルターを設定
    ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol)).AutoFilter
    
    ' 画面更新を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' 完了メッセージ
    'MsgBox "フィルターが設定されました。", vbInformation
End Sub


Sub RemoveFilter_G2_ErrorData()
    Dim ws As Worksheet


    Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")

    ' フィルターが設定されている場合、フィルターを解除
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
        ws.AutoFilterMode = False
    End If

    ' 完了メッセージ
    'MsgBox "フィルターが解除されました。", vbInformation
End Sub
