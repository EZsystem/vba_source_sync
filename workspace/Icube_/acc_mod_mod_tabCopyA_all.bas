Attribute VB_Name = "mod_tabCopyA_all"
'===============================================
' モジュール名 : mod_tabCopyA_cleanThenInsert
' 処理名       : mod_tabCopyA_cleanThenInsert_Execute
' 説明         : 転写元のキーに基づき転写先の行を削除した後、
'                重複していない転写元データのみを転写するにゃ
' 使用クラス   : acc_clsTableCleaner, acc_clsTableInserter
' 作成日       : 2025/05/09
'===============================================
Option Compare Database
Option Explicit

Public Sub mod_tabCopyA_cleanThenInsert_Execute()
    On Error GoTo EH

    ' --- 1. 転写先データの削除（Cleaner） ---
    Dim cleaner As New acc_clsTableCleaner
    cleaner.Init
    cleaner.AddSetting "Icube_", "kt_基本工事_完工", "基本工事コード"
    cleaner.AddSetting "Icube_", "kt_基本工事_作業所", "基本工事コード"
    cleaner.AddSetting "Icube_", "kt_基本工事_受注", "基本工事コード"
    cleaner.AddSetting "Icube_", "kt_工事コード情報", "工事コード"
    cleaner.AddSetting "Icube_", "kt_枝番工事", "枝番工事コード"
    cleaner.CleanTarget

    ' --- 2. 重複なしのレコードのみ転写（Inserter） ---
    Dim inserter As New acc_clsTableInserter
    inserter.Init
    inserter.AddSetting "Icube_", "kt_基本工事_完工", "基本工事コード"
    inserter.AddSetting "Icube_", "kt_基本工事_作業所", "基本工事コード"
    inserter.AddSetting "Icube_", "kt_基本工事_受注", "基本工事コード"
    inserter.AddSetting "Icube_", "kt_工事コード情報", "工事コード"
    inserter.AddSetting "Icube_", "kt_枝番工事", "枝番工事コード"
    inserter.InsertUniqueOnly

    MsgBox "削除＋重複除外付き転写が完了したにゃ！", vbInformation
    Exit Sub

EH:
    MsgBox "【エラー発生】：" & vbCrLf & Err.description, vbCritical
End Sub

