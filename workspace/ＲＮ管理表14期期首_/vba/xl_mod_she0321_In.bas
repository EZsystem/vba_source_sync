Attribute VB_Name = "she0321_In"
Option Explicit
'シートD2へファイルから取込み
Sub ImportDataFromUserFileD2()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim sourceBook As Workbook
    Dim sourceFilePath As Variant
    
    ' ユーザーにファイルを選択させる
    sourceFilePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "ファイルを選択してください")
    If sourceFilePath = False Then
        MsgBox "ファイルの選択がキャンセルされました。", vbExclamation
        Exit Sub
    End If
    
    ' 選択されたファイルを開く
    Set sourceBook = Workbooks.Open(sourceFilePath)
    Set wsSource = sourceBook.Sheets(1)
    
    ' 取り込み先シートの設定
    Set wsDest = ThisWorkbook.Sheets("D2損益期中")
    
    ' 画面更新、計算設定、警告表示を無効化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' フィルタ解除
    If wsSource.AutoFilterMode Then
        wsSource.AutoFilterMode = False
    End If
    If wsDest.AutoFilterMode Then
        wsDest.AutoFilterMode = False
    End If
    
    ' 取り込み先シートのデータをクリア
    wsDest.Cells.Clear
    
    ' シート全体のデータをコピー
    wsSource.UsedRange.Copy
    wsDest.Range("A5").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    wsDest.Range("A5").PasteSpecial Paste:=xlPasteFormats
    
    ' コピー後の設定をクリア
    Application.CutCopyMode = False
    
    ' 選択されたファイルを閉じる
    sourceBook.Close SaveChanges:=False
    
    ' 画面更新、計算設定、警告表示を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    MsgBox "データの取り込みが完了しました。", vbInformation
End Sub
