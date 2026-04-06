Attribute VB_Name = "acc_mod_KoguchiFinalizer"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : acc_mod_KoguchiFinalizer
' 概要           : 実績Aの3期平均を算出し、完工高予測のベースを作成するプログラム
' 更新日         : 2026/04/03
'===================================================================================================

' Private Const TARGET_TABLE As String = "at_Work_02_受注_3期平均"
' Private Const SRC_ACTUALS  As String = "at_Work_01_実績推移_3期分"
' Private Const SRC_FORECAST As String = "at_Work_04_受注_今期予測"

Private Const TARGET_TABLE As String = AT_WORK_02_ORDER_3P_AVE
Private Const SRC_ACTUALS  As String = AT_WORK_01_ACTUALS_3P
Private Const SRC_FORECAST As String = AT_WORK_04_ORDER_FCST

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Run_Final_Aggregation_Reset
' 概要           : 3期平均ロジックに基づき、完工高予測ベーステーブルを再構築します。
'---------------------------------------------------------------------------------------------------
Public Sub Run_Final_Aggregation_Reset(Optional isBatch As Boolean = False)
    On Error GoTo Err_Handler
    Dim db As DAO.Database: Set db = CurrentDb
    
    Debug.Print "--- 完工高予測3期平均 算出開始: " & Now & " ---"
    
    ' 1. テーブルの初期化
    Call Initialize_Target_Table(db)
    
    ' 2. データの集計と書き出し
    Call Aggregate_Final_Baseline_3YearAverage(db)
    
    Debug.Print "--- 集計が正常に完了しました ---"
    If Not isBatch Then
        MsgBox "テーブル [" & TARGET_TABLE & "] の作成が完了しました。", vbInformation, "完了"
    End If
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
    Dim sql As String
    sql = "CREATE TABLE [" & TARGET_TABLE & "] (" & _
          "ID AUTOINCREMENT PRIMARY KEY, " & _
          "期_ターゲット TEXT(20), " & _
          "受注月 TEXT(20), " & _
          "実績A_3期平均 CURRENCY)"
    db.Execute sql
End Sub

'---------------------------------------------------------------------------------------------------
' 2. 3期平均実績の集計
'---------------------------------------------------------------------------------------------------
Private Sub Aggregate_Final_Baseline_3YearAverage(ByRef db As DAO.Database)
    Dim rsIn As DAO.Recordset
    Dim rsOut As DAO.Recordset
    Dim dictA_Sum As Object: Set dictA_Sum = CreateObject("Scripting.Dictionary")
    
    ' ?@ 履歴テーブルから「3期分すべて」の全社合計をメモリにロード
    ' 条件: 受注期 = 完工期 (同一期内完工レコード)
    Dim sqlA As String
    sqlA = "SELECT [受注月_], Sum([工事価格の合計]) AS 月実績合計 " & _
           "FROM [" & SRC_ACTUALS & "] " & _
           "WHERE [受注期_] = [完工期_] " & _
           "GROUP BY [受注月_]"
    Set rsIn = db.OpenRecordset(sqlA, dbOpenSnapshot)
    Do While Not rsIn.EOF
        dictA_Sum.Add Nz(rsIn![受注月_]), Nz(rsIn![月実績合計], 0)
        rsIn.MoveNext
    Loop
    rsIn.Close
    
    ' ?A テーブルへの書き出し (4月始まり順)
    Set rsOut = db.OpenRecordset(TARGET_TABLE, dbOpenDynaset)
    Dim monthOrder As Variant: monthOrder = Array("4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月")
    
    ' ターゲット期の特定 (予測テーブルの最新期、または Icubeの最新期+1 など)
    ' ここでは予測テーブルにある最新ターゲット期を取得
    Dim targetP As String
    On Error Resume Next
    targetP = db.OpenRecordset("SELECT Max([予測ターゲット]) FROM [" & SRC_FORECAST & "]").Fields(0)
    If Err.Number <> 0 Then targetP = "次期" ' 予測テーブルが空の場合のフォールバック
    On Error GoTo 0
    
    Dim m As Variant
    For Each m In monthOrder
        Dim valASum As Currency: valASum = 0
        Dim valAAve As Currency: valAAve = 0
        
        If dictA_Sum.Exists(m) Then valASum = dictA_Sum(m)
        
        ' 実績平均 (3で割り、四捨五入して整数に)
        ' Currency型の演算で端数が出るため、Int(x + 0.5) で丸めます
        valAAve = Int((valASum / 3) + 0.5)
        
        rsOut.AddNew
        rsOut![期_ターゲット] = targetP
        rsOut![受注月] = m
        rsOut![実績A_3期平均] = valAAve
        rsOut.Update
    Next
    rsOut.Close
End Sub
