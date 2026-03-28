Attribute VB_Name = "she02_disp"
Option Explicit

'表示期首、期中
Sub she02_dispSet1()
    Dim ws As Worksheet
Call she02_dispReset
    ' シートID: 2 に対応するシートを取得
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets(2) ' シートID: 2
    
    ' 非表示設定
    ws.Rows("1:25").Hidden = True ' 1から25行を非表示
    ws.Columns("A:E").Hidden = True ' AからF列を非表示

    ' ウインドウ枠固定
    With ws
        .Activate ' ウインドウ枠固定のためにシートをアクティブ化
        ActiveWindow.FreezePanes = False ' 既存のウインドウ枠を解除
        .Cells(37, "I").Select ' I37セルを選択
        ActiveWindow.FreezePanes = True ' ウインドウ枠を固定
    End With

    'MsgBox "シート表示範囲の設定が完了しました。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
End Sub

'表示期中
Sub she02_dispSet2()
    Dim ws As Worksheet
Call she02_dispReset
    ' シートID: 2 に対応するシートを取得
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets(2) ' シートID: 2

    ' 非表示設定
    ws.Rows("1:25").Hidden = True ' 1から25行を非表示
    ws.Columns("A:AG").Hidden = True ' AからAG列を非表示

    ' ウインドウ枠固定
    With ws
        .Activate ' ウインドウ枠固定のためにシートをアクティブ化
        ActiveWindow.FreezePanes = False ' 既存のウインドウ枠を解除
        .Cells(37, "AK").Select ' AK37セルを選択
        ActiveWindow.FreezePanes = True ' ウインドウ枠を固定
    End With

    ' 完了メッセージ
    'MsgBox "シート表示範囲の設定が完了しました。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
End Sub

'表示設定解除
Sub she02_dispReset()
    Dim ws As Worksheet
    
    ' シートID: 2 に対応するシートを取得
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets(2) ' シートID: 2

    ' 行と列の再表示
    ws.Rows.Hidden = False ' 全ての行を再表示
    ws.Columns.Hidden = False ' 全ての列を再表示

    ' ウインドウ枠解除
    With ws
        .Activate ' ウインドウ枠解除のためにシートをアクティブ化
        ActiveWindow.FreezePanes = False ' ウインドウ枠を解除
    End With

    ' 完了メッセージ
    'MsgBox "シート全体の再表示とウインドウ枠の解除が完了しました。", vbInformation

    ' セル I32 を選択
    ws.Cells(32, "I").Select

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
End Sub

