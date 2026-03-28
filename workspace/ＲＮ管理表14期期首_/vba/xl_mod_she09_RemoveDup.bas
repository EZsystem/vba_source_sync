Attribute VB_Name = "she09_RemoveDup"
' ============================================
' Module   : she09_RemoveDup
' プロシージャ名 : RemoveDuplicates_ByColA
' 概要         : 工事一覧テーブルの A 列をキーに重複行を削除する
' 対象範囲     : テーブル tbl_工事一覧 のデータ範囲（A列～最終列）
' 条件         : A列が重複する場合、最初の1行を残して削除
' 備考         : ListObject.Range.RemoveDuplicates を使用
' ============================================

Sub RemoveDuplicates_ByColA()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim keyColumnIndex As Long

    ' --- 1. シート・テーブル設定 ---
    Set ws = ThisWorkbook.Sheets("tbl")
    Set lo = ws.ListObjects("tbl_工事一覧")

    ' A列（s基本工事コード）の ListObject 内での列番号
    ' → 列名で指定（列が移動しても壊れない）
    keyColumnIndex = lo.ListColumns("s基本工事コード").Index

    ' --- 2. 重複削除（テーブル全体に対して実行）---
    lo.Range.RemoveDuplicates Columns:=keyColumnIndex, Header:=xlYes

End Sub
