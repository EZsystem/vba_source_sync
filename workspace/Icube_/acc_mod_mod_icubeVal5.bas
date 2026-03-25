Attribute VB_Name = "mod_icubeVal5"
'-------------------------------------
' Module: mod_icubeVal5
' 説明　：acc_clsFieldTranscriber クラスを使用して、フィールド転写を実行する
' 作成日：2025/05/12
' 更新日：-
'-------------------------------------
Option Compare Database
Option Explicit

'=================================================
' サブルーチン名 : Run_FieldTranscribe_Default
' 説明   : Icube_ テーブルの「基本工事コード」「基本工事名称」を
'          対応する転写先フィールドにコピーする（Null値はスキップ）
'=================================================
Public Sub Run_FieldTranscribe_Default()
    On Error GoTo Err_Handler

    Dim trans As acc_clsFieldTranscriber
    Set trans = New acc_clsFieldTranscriber

    With trans
        .Init "Icube_", , True  ' テーブル名：Icube_、条件なし、Nullスキップ
        .AddMapping "基本工事コード", "s基本工事コード"
        .AddMapping "基本工事名称", "s基本工事名称"
        .TranscribeAll
    End With

    'MsgBox "フィールド転写が完了しましたにゃ！", vbInformation
    Exit Sub

Err_Handler:
    MsgBox "【MainUpdater】エラー：" & Err.description, vbCritical
    Debug.Print "【MainUpdater】Err:" & Err.Number & " - " & Err.description
End Sub


