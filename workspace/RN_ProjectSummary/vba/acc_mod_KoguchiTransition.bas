Attribute VB_Name = "acc_mod_KoguchiTransition"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : acc_mod_KoguchiTransition
' 概要           : 受注月→完工月の推移パターン（3期平均値・割合）を算出するプログラム
' 更新日         : 2026/04/03
'===================================================================================================

Private Const TARGET_TABLE  As String = "at_Work_小口完工推移3期平均"
Private Const SRC_HISTORY   As String = "at_Work_小口受注完工推移_3期分"
Private Const SRC_FORECAST  As String = "at_Work_受注完工予測_加重平均集計"
Private Const SRC_BASELINE  As String = "at_Work_完工高予測3期平均"

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Run_Transition_Aggregation_Reset
' 概要           : 実績推移データから、受注・完工月の組み合わせ別の3期平均値と割合を算出します。
'---------------------------------------------------------------------------------------------------
Public Sub Run_Transition_Aggregation_Reset()
    On Error GoTo Err_Handler
    Dim db As DAO.Database: Set db = CurrentDb
    
    Debug.Print "--- 受注完工推移パターン(3期平均) 算出開始: " & Now & " ---"
    
    ' 1. テーブルの初期化
    Call Initialize_Target_Table(db)
    
    ' 2. データの集計と書き出し
    Call Aggregate_Transition_Ratio_3YearAverage(db)
    
    Debug.Print "--- 集計が正常に完了しました ---"
    MsgBox "テーブル [" & TARGET_TABLE & "] の作成が完了しました。", vbInformation, "完了"
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
          "期_予測ターゲット TEXT(20), " & _
          "受注月 TEXT(20), " & _
          "完工月 TEXT(20), " & _
          "3期平均値 CURRENCY, " & _
          "完工高割合 DOUBLE)"
    db.Execute sql
End Sub

'---------------------------------------------------------------------------------------------------
' 2. 推移パターンの平均および割合の算出
'---------------------------------------------------------------------------------------------------
Private Sub Aggregate_Transition_Ratio_3YearAverage(ByRef db As DAO.Database)
    Dim rsBase As DAO.Recordset
    Dim rsIn As DAO.Recordset
    Dim rsOut As DAO.Recordset
    Dim dictBaseline As Object: Set dictBaseline = CreateObject("Scripting.Dictionary")
    Dim dictAgg As Object: Set dictAgg = CreateObject("Scripting.Dictionary")
    
    ' �@ 分母となる「受注月ごとの3期平均実績」をメモリにロード
    Set rsBase = db.OpenRecordset("SELECT [受注月], [実績A_3期平均] FROM [" & SRC_BASELINE & "]", dbOpenSnapshot)
    Do While Not rsBase.EOF
        dictBaseline.Add Nz(rsBase![受注月]), Nz(rsBase![実績A_3期平均], 0)
        rsBase.MoveNext
    Loop
    rsBase.Close
    
    ' �A 履歴テーブルから「受注月×完工月」の全社合計をメモリにロード
    ' 条件: 受注期 = 完工期
    Dim sqlIn As String
    sqlIn = "SELECT [受注月_], [完工月_], Sum([工事価格の合計]) AS ペア実績合計 " & _
            "FROM [" & SRC_HISTORY & "] " & _
            "WHERE [受注期_] = [完工期_] " & _
            "GROUP BY [受注月_], [完工月_]"
    Set rsIn = db.OpenRecordset(sqlIn, dbOpenSnapshot)
    Do While Not rsIn.EOF
        Dim key As String: key = Nz(rsIn![受注月_]) & "|" & Nz(rsIn![完工月_])
        dictAgg.Add key, Nz(rsIn![ペア実績合計], 0)
        rsIn.MoveNext
    Loop
    rsIn.Close
    
    ' �B テーブルへの書き出し
    Set rsOut = db.OpenRecordset(TARGET_TABLE, dbOpenDynaset)
    
    ' ターゲット期の特定
    Dim targetP As String
    On Error Resume Next
    targetP = db.OpenRecordset("SELECT Max([予測ターゲット]) FROM [" & SRC_FORECAST & "]").Fields(0)
    If Err.Number <> 0 Then targetP = "次期"
    On Error GoTo 0
    
    Dim k As Variant, parts As Variant
    For Each k In dictAgg.Keys
        parts = Split(k, "|")
        Dim sMonth As String: sMonth = parts(0)
        
        Dim valSum3Period As Currency: valSum3Period = dictAgg(k)
        Dim valAve3Period As Currency: valAve3Period = Int((valSum3Period / 3) + 0.5)
        
        ' 完工高割合の算出 (平均期待値 / 受注月のベース平均)
        Dim ratio As Double: ratio = 0
        If dictBaseline.Exists(sMonth) Then
            Dim baseVal As Currency: baseVal = dictBaseline(sMonth)
            If baseVal <> 0 Then
                ratio = CDbl(valAve3Period / baseVal)
            End If
        End If
        
        rsOut.AddNew
        rsOut![期_予測ターゲット] = targetP
        rsOut![受注月] = sMonth
        rsOut![完工月] = parts(1)
        rsOut![3期平均値] = valAve3Period
        rsOut![完工高割合] = ratio
        rsOut.Update
    Next
    rsOut.Close
    
    Debug.Print "書き出し件数: " & dictAgg.count & "件"
End Sub
