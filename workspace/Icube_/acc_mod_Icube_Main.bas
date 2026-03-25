Attribute VB_Name = "acc_mod_Icube_Main"
'Attribute VB_Name = "acc_mod_Icube_Main"
Option Compare Database
Option Explicit

'===================================================================================================
' プロシージャ名 : Run_All_iCubeValidation_WithLog
' 概要           : これが実行の入り口です。F5キーでこれを動かしてください。
'===================================================================================================
Public Sub Run_All_iCubeValidation_WithLog()
    Dim clsLog As com_clsErrorUtility
    Dim clsCleaner As acc_clsDataCleaner
    Dim clsTransfer As acc_clsIcubeTransfer  ' ★クラスを追加
    
    Set clsLog = New com_clsErrorUtility
    Set clsCleaner = New acc_clsDataCleaner
    Set clsTransfer = New acc_clsIcubeTransfer ' ★インスタンス生成
    
    clsLog.Init isBatch:=True
    clsCleaner.Init
    clsTransfer.Init clsLog ' ★クラスの初期化
    
    On Error GoTo Err_Handler
    
    Call ClearLog
    Call AppendLog("--- iCube一括処理 開始 ---")

    ' 1. バリデーション (Validatorモジュール呼び出し)
    Call AppendLog(">> 工程1: バリデーション実行")
    Call Process_BasicValidation_And_Split(clsCleaner, clsLog)
    Call Process_Category_And_Price(clsLog)
    Call Process_Transcribe_ProjectInfo(clsLog)
    Call Process_Final_Cleansing(clsLog)

    ' 2. データ転写 (クラスのメソッドを呼び出し)
    Call AppendLog(">> 工程2: データ統合・関連転写実行")
    clsTransfer.ToHistory      ' ★メソッド呼び出し
    clsTransfer.ToRelatedTables ' ★メソッド呼び出し

    Call AppendLog("--- 全工程 正常終了 ---")
    clsLog.Show_Final_Report
    
    Exit Sub
Err_Handler:
    clsLog.Notify_Smart_Popup "Main Error: " & Err.Description
End Sub
