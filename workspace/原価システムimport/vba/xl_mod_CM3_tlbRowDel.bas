Attribute VB_Name = "CM3_tlbRowDel"
'-------------------------------------
' Module: CM3_tlbRowDel
' 説明　：テーブル操作の共通関数群
' 作成日：2025/12/01
' 更新日：-
'-------------------------------------

Option Explicit

'============================================
' プロシージャ名         : Sub Delete_TableRecordsFromSecondRow
' 概要                   : 指定テーブルの2行目以降のレコードを高速削除する
' 引数                   : なし
' 戻り値（Functionのみ） : -
' 呼び出し元フォーム／イベント : 未特定
' 関連情報               : シート「tbl」のテーブル「tbl_工事一覧」を対象とする
' 備考                   : 範囲一括削除により高速処理を実現
'============================================
Public Sub Delete_TableRecordsFromSecondRow()
    
    ' --- 1. 初期化 ---
    Dim targetSheet As Worksheet
    Dim targetTable As ListObject
    Dim dataRowCount As Long
    Dim deleteRange As Range
    
    ' 画面更新を停止して処理速度を向上
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- 2. 条件判定 ---
    ' シートの存在確認
    Set targetSheet = ThisWorkbook.Worksheets("原価S_基本工事")
    If targetSheet Is Nothing Then
        MsgBox "シート「tbl」が見つかりません", vbCritical
        GoTo CleanUp
    End If
    
    ' テーブルの存在確認
    Set targetTable = targetSheet.ListObjects("tbl_原価S_基本工事")
    If targetTable Is Nothing Then
        MsgBox "テーブル「tbl_原価S_基本工事」が見つかりません", vbCritical
        GoTo CleanUp
    End If
    
    ' データ行数の確認
    dataRowCount = targetTable.ListRows.Count
    If dataRowCount <= 0 Then
        MsgBox "削除対象のデータ行がありません", vbInformation
        GoTo CleanUp
    End If
    
    ' --- 3. 実行処理 ---
    ' データ範囲を一括で取得
    Set deleteRange = targetTable.DataBodyRange
    
    ' ※注意：範囲全体を一括削除することで高速化を実現
    If Not deleteRange Is Nothing Then
        deleteRange.Delete Shift:=xlUp
    End If
    
    ' --- 4. 結果の出力 ---
    'MsgBox "テーブル「tbl_原価S_基本工事」の全データ行を削除しました（" & dataRowCount & "行削除）", vbInformation
    
CleanUp:
    ' 画面更新と計算を復元
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub   ' ← Subの終わり

