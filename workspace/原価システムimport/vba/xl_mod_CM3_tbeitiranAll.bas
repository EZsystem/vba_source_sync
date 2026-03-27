Attribute VB_Name = "CM3_tbeitiranAll"

' Module: CM3_tbeitiranAll
' プロシージャ名 : she09_tbeitiranAll_1

'工事データ更新
Sub CM3_tbeitiranAll_1()

'テーブルデータ(tbl_工事一覧)のクリア
    Call Delete_TableRecordsFromSecondRow
'Accessから一件工事データ転写
    Call Inport_koujiitiran_FromAccess
    

MsgBox "Accessからデータを取得してExcelに出力しました。", vbInformation
End Sub
