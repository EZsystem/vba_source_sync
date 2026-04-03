Attribute VB_Name = "acc_mod_KoguchiFinalForecast"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : acc_mod_KoguchiFinalForecast
' 概要           : 予測受注額と実績完工割合を統合し、最終的な予測完工高を算出するプログラム
' 更新日         : 2026/04/03
'===================================================================================================

Private Const TARGET_TABLE  As String = "at_Work_予測完工高_最終結果"
Private Const SRC_FORECAST  As String = "at_Work_受注完工予測_加重平均集計"
Private Const SRC_PATTERN   As String = "at_Work_小口完工推移3期平均"

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Run_Final_Forecast_Reset
' 概要           : 加重平均予測受注額に完工割合を適用し、最終的な完工高予測テーブルを作成します。
'---------------------------------------------------------------------------------------------------
Public Sub Run_Final_Forecast_Reset()
    On Error GoTo Err_Handler
    Dim db As DAO.Database: Set db = CurrentDb
    
    Debug.Print "--- 最終完工高予測 算出開始: " & Now & " ---"
    
    ' 1. テーブルの初期化
    Call Initialize_Target_Table(db)
    
    ' 2. データの集計と書き出し
    Call Aggregate_Final_Forecast_To_Table(db)
    
    Debug.Print "--- すべての工程が正常に完了しました ---"
    MsgBox "最終予測テーブル [" & TARGET_TABLE & "] の作成が完了しました。", vbInformation, "完了"
    Exit Sub

Err_Handler:
    MsgBox "算出エラー: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------------------------------------
' 1. テーブル初期化
'---------------------------------------------------------------------------------------------------
Private Sub Initialize_Target_Table(ByRef db As DAO.Database)
    On Error Resume Next
    db.Execute "DROP TABLE [" & TARGET_TABLE & "]"
    Err.Clear
    On Error GoTo 0
    
    Debug.Print "テーブルを新規作成中: " & TARGET_TABLE
    ' ご提示いただいた項目 + 施工管轄組織名 + 検証用証跡カラム
    Dim sql As String
    sql = "CREATE TABLE [" & TARGET_TABLE & "] (" & _
          "ID AUTOINCREMENT PRIMARY KEY, " & _
          "期_予測ターゲット TEXT(20), " & _
          "施工管轄組織名 TEXT(100), " & _
          "受注月 TEXT(20), " & _
          "完工月 TEXT(20), " & _
          "元_受注予測額 CURRENCY, " & _
          "適用比率 DOUBLE, " & _
          "予測完工高 CURRENCY)"
    db.Execute sql
End Sub

'---------------------------------------------------------------------------------------------------
' 2. 完工高予測の集計
'---------------------------------------------------------------------------------------------------
Private Sub Aggregate_Final_Forecast_To_Table(ByRef db As DAO.Database)
    Dim rsFor As DAO.Recordset
    Dim rsPat As DAO.Recordset
    Dim rsOut As DAO.Recordset
    Dim dictPatterns As Object: Set dictPatterns = CreateObject("Scripting.Dictionary")
    
    ' �@ 完工パターン割合をメモリにロード (Dictionary in Dictionary)
    ' 構造: 受注月 -> { 完工月1: 割合1, 完工月2: 割合2, ... }
    Set rsPat = db.OpenRecordset("SELECT [受注月], [完工月], [完工高割合] FROM [" & SRC_PATTERN & "]", dbOpenSnapshot)
    Do While Not rsPat.EOF
        Dim om As String: om = Nz(rsPat![受注月])
        Dim cm As String: cm = Nz(rsPat![完工月])
        Dim rt As Double: rt = CDbl(Nz(rsPat![完工高割合], 0))
        
        If Not dictPatterns.Exists(om) Then
            dictPatterns.Add om, CreateObject("Scripting.Dictionary")
        End If
        dictPatterns(om).Add cm, rt
        rsPat.MoveNext
    Loop
    rsPat.Close
    
    ' �A 予測受注額(ベース)をループして、パターンを適用して書き出す
    Set rsFor = db.OpenRecordset("SELECT * FROM [" & SRC_FORECAST & "]", dbOpenSnapshot)
    Set rsOut = db.OpenRecordset(TARGET_TABLE, dbOpenDynaset)
    
    Do While Not rsFor.EOF
        Dim orderMonth As String: orderMonth = Nz(rsFor![受注月])
        Dim weightedOrderAmt As Currency: weightedOrderAmt = Nz(rsFor![加重集計値], 0)
        
        ' 該当する受注月の完工パターンがあるか確認
        If dictPatterns.Exists(orderMonth) Then
            Dim dictCM As Object: Set dictCM = dictPatterns(orderMonth)
            Dim finishMonth As Variant
            
            ' すべての完工月候補に対して計算
            For Each finishMonth In dictCM.Keys
                Dim ratio As Double: ratio = dictCM(finishMonth)
                Dim predictedAmt  As Currency
                
                ' 予測完工高 = 受注予測額 × 完工高割合 (1円単位で四捨五入)
                predictedAmt = Int((weightedOrderAmt * ratio) + 0.5)
                
                ' 書き出し
                rsOut.AddNew
                rsOut![期_予測ターゲット] = rsFor![予測ターゲット]
                rsOut![施工管轄組織名] = rsFor![施工管轄組織名]
                rsOut![受注月] = orderMonth
                rsOut![完工月] = finishMonth
                ' --- 証跡の追加 ---
                rsOut![元_受注予測額] = weightedOrderAmt
                rsOut![適用比率] = ratio
                ' ------------------
                rsOut![予測完工高] = predictedAmt
                rsOut.Update
            Next
        End If
        rsFor.MoveNext
    Loop
    
    rsFor.Close: rsOut.Close
End Sub
