Attribute VB_Name = "mod_icubeAll"
'-------------------------------------
' Module: acc_mod_IcubeProcess
' 説明  : iCubeデータ一括処理フロー実行用モジュール
' 作成日: 2025/06/27
' 更新日: -
'-------------------------------------
Option Compare Database
Option Explicit

'============================================
' プロシージャ名 : Execute_Icube_FullSequence
' Module         : acc_mod_IcubeProcess
' 概要           : データインポート・値加工・累計テーブル転写を一括で実行する
' 引数           : なし
' 戻り値         : なし
' 呼び出し元     : 任意（フォームボタン・他モジュール等）
' 関連情報       : iCube運用の主要処理を連続自動化するためのラッパー
' 備考           : 各プロシージャはエラー発生時に即中断
'============================================
Public Sub Execute_Icube_FullSequence()
    On Error GoTo ErrHandler

    ' --- 1. データインポート処理 ---
    ' ExcelやCSV等からデータを仮テーブル等へ取り込む
    Call Run_IcubeImport_FullSequence

    ' --- 2. データ値加工処理 ---
    ' 必要な変換・クリーニング・検証を実施（ログ出力あり）
    Call Run_All_iCubeValidation_WithLog

    ' --- 3. 累計テーブル転写処理 ---
    ' 仮テーブルから本テーブル等へ転写（INSERT等）
    Call Transfer_IcubeData

    Exit Sub

ErrHandler:
    ' いずれかの処理でエラーが発生した場合は即中断
    MsgBox "エラー発生のため処理を中断しました。" & vbCrLf & _
           "処理名: Execute_Icube_FullSequence" & vbCrLf & _
           "エラー内容: " & Err.Number & " / " & Err.description, vbCritical
    Stop
End Sub


