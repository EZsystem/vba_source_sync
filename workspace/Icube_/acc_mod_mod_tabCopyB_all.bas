Attribute VB_Name = "mod_tabCopyB_all"
Option Compare Database
Option Explicit
'テストしてないよ　2025/05/07　16:03
'-----------------------------------------------------
' モジュール名: mod_tabCopy2Ball
' 処理名　　: mod_tabCopy2B_all
' 説明　　　: Icube_累計 から関連テーブルへ一括転写
'           : 事前に転写先テーブルを全件削除し、重複削除→転写を行うにゃ
' 使用クラス : clsTableTransferSetting
' 作成日　　: 2025/05/09
'-----------------------------------------------------
Public Sub mod_tabCopy2B_all()
    Dim transferList As Collection
    Set transferList = New Collection

    Dim setting As clsTableTransferSetting

    ' --- 転写設定の登録 ---
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.targetTable = "kt_基本工事_完工"
    setting.keyField = "基本工事コード"
    transferList.Add setting

    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.targetTable = "kt_基本工事_作業所"
    setting.keyField = "基本工事コード"
    transferList.Add setting

    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.targetTable = "kt_基本工事_受注"
    setting.keyField = "基本工事コード"
    transferList.Add setting

    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.targetTable = "kt_工事コード情報"
    setting.keyField = "工事コード"
    transferList.Add setting

    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.targetTable = "kt_枝番工事"
    setting.keyField = "枝番工事コード"
    transferList.Add setting

    ' --- 実行処理 ---
    Dim settingItem As clsTableTransferSetting
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim sqlDelete As String
    Set db = CurrentDb

    For Each settingItem In transferList
        Debug.Print "▼ 転写準備：[" & settingItem.targetTable & "] を全件削除中..."
        db.Execute "DELETE FROM [" & settingItem.targetTable & "]", dbFailOnError
        Debug.Print "　→ 全件削除完了"

        Debug.Print "▼ 転写処理中: [" & settingItem.SourceTable & "] → [" & settingItem.targetTable & "]"

        ' --- 重複データ削除 ---
        Set rsSource = db.OpenRecordset("SELECT DISTINCT [" & settingItem.keyField & "] FROM [" & settingItem.SourceTable & "]", dbOpenSnapshot)
        Do While Not rsSource.EOF
            sqlDelete = "DELETE FROM [" & settingItem.targetTable & "] " & _
                        "WHERE [" & settingItem.keyField & "] = '" & replace(rsSource(settingItem.keyField), "'", "''") & "'"
            db.Execute sqlDelete, dbFailOnError
            rsSource.MoveNext
        Loop
        rsSource.Close
        Set rsSource = Nothing

        ' --- データ転写 ---
        TransferTable settingItem.SourceTable, settingItem.targetTable, settingItem.keyField
    Next

    Debug.Print "■■ 転写処理がすべて完了したにゃ！ [" & Now & "] ■■"
    'MsgBox "転写処理が完了したにゃ！", vbInformation

    Set db = Nothing
End Sub


