Attribute VB_Name = "acc_mod_Import_Staff"
Option Explicit

'--------------------------------------------
' プロシージャ名： Import_From_Bridge_Form
' 概要： フォームからデータを取得し、テーブルを初期化した後に登録する
'--------------------------------------------

'-------------------------------------
' Module: acc_mod_Import_Staff (Refactored)
' 必要コンポーネント:
'   - acc_clsTableUpdater (マスタ同期)
'   - com_mod_StringUtilities (正規化)
'   - acc_clsTransactionManager (データ整合性)
'-------------------------------------

Public Sub Import_Staff_EZ()
    Dim updater As New acc_clsTableUpdater
    Dim trans As New acc_clsTransactionManager
    Dim rawData As String
    
    ' 1. バッファ取得
    rawData = Forms("frm_Staff_Import_Bridge").Controls("txt_DataBuffer").value
    If Trim(rawData) = "" Then Exit Sub

    Call Fast_Mode_Toggle(True)
    
    On Error GoTo ErrLine
    updater.Init ' 初期化を実行
    trans.BeginTransaction ' 正しいメソッド名
    
    ' 3. 同期設定と実行
    With updater
        .TargetTable = "at_社員情報"
        .TempTable = "at_社員情報Temp"
        .KeyField = "社員番号"
        .FieldMapping = Array("氏名_戸籍上", "氏名カナ", "氏名_ﾒｰﾙ表示用", "資格", "所属", "役職", "対外呼称")
        
        If .ImportFromBuffer(rawData, vbCrLf, vbTab) Then
            .SyncWithMaster
        End If
    End With
    
    ' ★ 修正: CommitTrans ではなく Commit (TransactionManagerの定義に合わせる)
    trans.Commit
    Call Notify_Smart_Popup("社員情報の同期が完了しました。", 3)

CleanUp:
    Call Fast_Mode_Toggle(False)
    Exit Sub

ErrLine:
    ' ★ 修正: RollbackTrans ではなく Rollback
    trans.Rollback
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
