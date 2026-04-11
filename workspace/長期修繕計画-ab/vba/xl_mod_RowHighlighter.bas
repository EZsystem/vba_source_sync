Attribute VB_Name = "xl_mod_RowHighlighter"

'-------------------------------------
' Module: xl_mod_RowHighlighter
' 説明  ：M列の特定値（行削除）を条件とした行の着色処理
' 構成  ：初期化 -> 検証 -> 実行 -> 終了処理 [cite: 11]
'-------------------------------------
Option Explicit ' [cite: 9]

' マジックナンバーの定義（ルール 15）
Private Const COLOR_YELLOW As Long = 65535 ' vbYellow
Private Const TARGET_VALUE As String = "行削除"
Private Const COL_M As String = "M"
Private Const START_COL As String = "B"
Private Const END_COL As String = "M"
Private Const START_ROW As Long = 2 ' 1行目を見出しと想定

''' <summary>
''' M列が「行削除」の場合、その行のB列からM列までを黄色く塗る
''' </summary>
Public Sub Highlight_DeleteRows()
    Dim shActive As Worksheet: Set shActive = ActiveSheet
    Dim lastRow As Long
    Dim r As Long
    Dim valM As String
    
    ' 1. 初期化・検証 [cite: 11]
    lastRow = shActive.Cells(shActive.Rows.count, COL_M).End(xlUp).Row
    If lastRow < START_ROW Then Exit Sub


    ' 2. 実行：高速モード開始
    Call Fast_Mode_Toggle(True)

    ' 前処理：既存の着色をクリアする場合はここに追加（今回は指定がないため維持）

    ' 3. ループ処理 [cite: 11]
    For r = START_ROW To lastRow
        ' M列の値を取得し、共通関数で正規化（スペース除去・半角化）
        valM = Normalize_Text(CStr(shActive.Cells(r, COL_M).value))
        
        ' 条件判定：正規化した値が「行削除」と一致するか
        If valM = Normalize_Text(TARGET_VALUE) Then
            ' 指定範囲（B列〜M列）を黄色に着色
            shActive.Range(shActive.Cells(r, START_COL), shActive.Cells(r, END_COL)).Interior.Color = COLOR_YELLOW
        End If
    Next r

    ' 4. 出力・終了処理 [cite: 11]
    Call Fast_Mode_Toggle(False) '
    Call Notify_Smart_Popup("「行削除」対象行の着色が完了しました。", "完了") ' [cite: 19]

End Sub
' ← プロシージャの終わり [cite: 12]

