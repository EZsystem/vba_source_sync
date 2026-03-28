Attribute VB_Name = "sheet03_filter"
Option Explicit

' このコードは、シート "G3_原価Sエラー調査" のフィルターを解除します。
' フィルターが設定されていても、いなくても、問題ないようにします。

Sub RemoveFilter_G3_ErrorData()
    Dim ws As Worksheet

    ' シート "G3_原価Sエラー調査" を設定
    Set ws = ThisWorkbook.Sheets("G3_原価Sエラー調査")

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



' このコードは、シート "G3_原価Sエラー調査" にフィルターを設定します。
' 検索行は D列に "小規模工事名" と書かれた行数です。
' 検索行から A列の最終行までの動的な範囲にフィルターを設定します。
' フィルターが掛かっていても、掛かっていなくても問題なく動作します。

Sub sheet03_filtering()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim searchRow As Long
    Dim lastCol As Long
    Dim rng As Range

    ' シート "G3_原価Sエラー調査" を設定
    Set ws = ThisWorkbook.Sheets("G3_原価Sエラー調査")

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' "小規模工事名" を含む行を検索
    searchRow = ws.Columns("D").Find(What:="小規模工事名", LookIn:=xlValues, LookAt:=xlWhole).row

    ' 最終列を取得
    lastCol = ws.Cells(searchRow, ws.Columns.Count).End(xlToLeft).Column

    ' フィルター範囲を設定
    Set rng = ws.Range(ws.Cells(searchRow, 1), ws.Cells(lastRow, lastCol))

    ' フィルターが設定されている場合は解除
    If ws.AutoFilterMode Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
        ws.AutoFilterMode = False
    End If

    ' フィルターを設定
    rng.AutoFilter

    ' 完了メッセージ
    'MsgBox "フィルターが設定されました。", vbInformation
End Sub
