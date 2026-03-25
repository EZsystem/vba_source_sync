Attribute VB_Name = "acc_mod_Icube_Main"
'Attribute VB_Name = "acc_mod_Icube_Main"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : acc_mod_Icube_Main
' 概要           : iCubeデータバリデーション・統合処理 メイン制御モジュール
' 依存コンポーネント:
'   - クラス     : com_clsErrorUtility (共通エラー/ログ管理)
'   - クラス     : acc_clsDataCleaner (Accessデータ洗浄)
'   - クラス     : acc_clsIcubeTransfer (データ転写ロジック)
'   - モジュール : acc_mod_Icube_Validator (バリデーション工程)
' 最終更新日     : 2026/03/26
'===================================================================================================

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Run_All_iCubeValidation_WithLog
' 概要           : iCube一括処理の全工程（洗浄・検証・転写）を統括するエントリポイントです。
' 引数           : なし
' 戻り値         : なし
' 備考           : 開発時または運用時は、本プロシージャを呼び出して実行してください。
'---------------------------------------------------------------------------------------------------
Public Sub Run_All_iCubeValidation_WithLog()
    ' --- 1. オブジェクト宣言 ---
    Dim clsLog As com_clsErrorUtility
    Dim clsCleaner As acc_clsDataCleaner
    Dim clsTransfer As acc_clsIcubeTransfer
    
    ' --- 2. インスタンス生成と初期化 ---
    Set clsLog = New com_clsErrorUtility
    Set clsCleaner = New acc_clsDataCleaner
    Set clsTransfer = New acc_clsIcubeTransfer
    
    ' 各クラスの初期化（ログクラスを各インスタンスへ伝搬させ、ログの一貫性を確保）
    clsLog.Init isBatch:=True
    clsCleaner.Init
    clsTransfer.Init clsLog
    
    ' エラーハンドリングの開始
    On Error GoTo Err_Handler
    
    ' --- 3. 処理開始ログ記録 ---
    Call ClearLog
    Call AppendLog("--- iCube一括処理 開始 ---")

    ' --- 4. 工程1: バリデーション (Validatorモジュール呼び出し) ---
    ' 各種ビジネスルールに基づいたデータの整合性チェックと分割処理を実行
    Call AppendLog(">> 工程1: バリデーション実行")
    Call Process_BasicValidation_And_Split(clsCleaner, clsLog)
    Call Process_Category_And_Price(clsLog)
    Call Process_Transcribe_ProjectInfo(clsLog)
    Call Process_Final_Cleansing(clsLog)

    ' --- 5. 工程2: データ転写 (Transferクラスのメソッドを呼び出し) ---
    ' 検証済みデータを履歴テーブルおよび関連マスタへ統合・反映
    Call AppendLog(">> 工程2: データ統合・関連転写実行")
    clsTransfer.ToHistory       ' 履歴テーブル(at_History)等への転写
    clsTransfer.ToRelatedTables  ' 関連するサブテーブルへの展開
    
    ' --- 6. 正常終了処理 ---
    Call AppendLog("--- 全工程 正常終了 ---")
    
    ' 累積されたログの最終レポートを表示
    clsLog.Show_Final_Report
    
    Exit Sub

Err_Handler:
    ' 予期せぬエラー発生時、ユーザーへ通知を行いログを保護
    clsLog.Notify_Smart_Popup "Main Error: " & Err.Description
End Sub
