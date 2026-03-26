Attribute VB_Name = "she09_tbeitiranAll"
' Module: she09_tbeitiranAll
' プロシージャ名 : she09_tbeitiranAll_1

'工事データ更新
Sub she09_tbeitiranAll_1()
'シート：工事一覧　が処理対象
'テーブルデータ(tbl_工事一覧)のクリア
    Call Delete_TableRecordsFromSecondRow
'Accessから一件工事データ転写
    Call Inport_koujiitiran_FromAccess
'テーブル重複の削除　対象列：s基本工事コード
    Call RemoveDuplicates_ByColA
    

MsgBox "Accessからデータを取得してExcelに出力しました。", vbInformation
End Sub


