Attribute VB_Name = "acc_mod_KoguchiMainProcessor"
Option Compare Database
Option Explicit

' =====================================
' Module: acc_mod_KoguchiMainProcessor
' 説明  : 小口工事予測の全工程（01-05）を一括実行する
' 作成日: 2026/04/06
' = : acc_mod_MappingTemplate (Public Consts)
' =====================================

''' <summary>
''' 小口工事予測の全ステップを順序通りに一括実行する。
''' 各ステップでのメッセージボックスは抑制され、最後に1回だけ完了通知を表示します。
''' </summary>
Public Sub Run_All_Koguchi_Calculations()
    Dim startTime As Date: startTime = Now
    Dim logMsg As String
    
    On Error GoTo Err_Handler
    
    Debug.Print "========== 小口工事予測一括処理 開始 [" & startTime & "] =========="
    
    ' ステップ 01: 実績推移の集計
    Debug.Print "[Step 01/05] 実績推移集計中 (" & AT_WORK_01_ACTUALS_3P & ")..."
    Call acc_mod_KoguchiAggregator.Run_Aggregator_Reset(isBatch:=True)
    
    ' ステップ 02: 3期平均（分母）の算出
    Debug.Print "[Step 02/05] 3期平均算出中 (" & AT_WORK_02_ORDER_3P_AVE & ")..."
    Call acc_mod_KoguchiFinalizer.Run_Final_Aggregation_Reset(isBatch:=True)
    
    ' ステップ 03: 推移割合（パターン）の算出
    Debug.Print "[Step 03/05] 推移割合算出中 (" & AT_WORK_03_COMP_RATIO & ")..."
    Call acc_mod_KoguchiTransition.Run_Transition_Aggregation_Reset(isBatch:=True)
    
    ' ステップ 04: 今期受注予測の計算
    Debug.Print "[Step 04/05] 今期受注予測計算中 (" & AT_WORK_04_ORDER_FCST & ")..."
    Call acc_mod_KoguchiForecast.Run_Aggregator_Weighted_Reset(isBatch:=True)
    
    ' ステップ 05: 最終完工予測の統合
    Debug.Print "[Step 05/06] 最終完工予測統合中 (" & AT_WORK_05_COMP_FCST & ")..."
    Call acc_mod_KoguchiFinalForecast.Run_Final_Forecast_Reset(isBatch:=True)
    
    ' ステップ 06: 給与・経費＋予測の最終集計
    Debug.Print "[Step 06/06] 給与・経費・予測の最終集計中 (" & AT_WORK_FINAL_AGGREGATION & ")..."
    Call acc_mod_KoguchiAggregationFinalizer.Run_KoguchiStaffExpense_Aggregation(isBatch:=True)
    
    Debug.Print "========== 一括処理 正常完了 [" & Now & "] (処理時間: " & Format(Now - startTime, "nn分ss秒") & ") =========="
    
    MsgBox "小口工事予測の全工程および最終集計（01?06）が正常に完了しました。" & vbCrLf & _
           "処理時間: " & Format(Now - startTime, "nn分ss秒"), vbInformation, "一括処理完了"
    Exit Sub

Err_Handler:
    logMsg = "ステップ実行中にエラーが発生しました。" & vbCrLf & _
             "エラー番号: " & Err.Number & vbCrLf & _
             "エラー内容: " & Err.Description
    Debug.Print "!!! エラー発生: " & logMsg
    MsgBox logMsg, vbCritical, "一括処理中断"
End Sub
