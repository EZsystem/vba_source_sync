Attribute VB_Name = "acc_mod_KoguchiForecast"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : acc_mod_KoguchiForecast
' 概要           : 加重平均集計テーブル (5項目) を作成・更新するプログラム
' 更新日         : 2026/04/03
'===================================================================================================

' Private Const TARGET_TABLE As String = "at_Work_04_受注_今期予測"
Private Const TARGET_TABLE As String = AT_WORK_04_ORDER_FCST

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Run_Aggregator_Weighted_Reset
' 概要           : 加重平均集計テーブルを係数テーブルに基づいて再構築します。
'---------------------------------------------------------------------------------------------------
Public Sub Run_Aggregator_Weighted_Reset(Optional isBatch As Boolean = False)
    On Error GoTo Err_Handler
    Dim db As DAO.Database: Set db = CurrentDb
    
    Debug.Print "--- 加重平均集計開始: " & Now & " ---"
    
    ' 1. テーブルの初期化
    Call Initialize_Target_Table(db)
    
    ' 2. 加重平均データの集計
    Call Aggregate_Weighted_To_Table(db)
    
    Debug.Print "--- 集計が正常に完了しました ---"
    If Not isBatch Then
        MsgBox "テーブル [" & TARGET_TABLE & "] の作成が完了しました。" & vbCrLf & _
               "内容をご確認ください。", vbInformation, "完了"
    End If
    Exit Sub

Err_Handler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "集計エラー"
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
          "予測ターゲット TEXT(20), " & _
          "施工管轄組織名 TEXT(100), " & _
          "Q TEXT(10), " & _
          "受注月 TEXT(10), " & _
          "加重集計値 CURRENCY)"
    db.Execute sql
End Sub

'---------------------------------------------------------------------------------------------------
' 2. 加重平均データの集計
'---------------------------------------------------------------------------------------------------
Private Sub Aggregate_Weighted_To_Table(ByRef db As DAO.Database)
    Dim rsKeisu As DAO.Recordset
    Dim rsIn As DAO.Recordset
    Dim rsOut As DAO.Recordset
    Dim dictAgg As Object: Set dictAgg = CreateObject("Scripting.Dictionary")
    
    ' 1. 係数マスタを読み込み
    Set rsKeisu = db.OpenRecordset("SELECT [期_計算対象], [作業所名], [加重率], Val(Replace(Nz([期_荷重対象],""""), ""期"", """")) AS 荷重期数値 FROM [at_受注額予測計数]", dbOpenSnapshot)
    
    ' 2. メモリ集計
    Do While Not rsKeisu.EOF
        Dim targetPeriod As String: targetPeriod = Nz(rsKeisu![期_計算対象])
        Dim shopName As String: shopName = Nz(rsKeisu![作業所名])
        Dim weightRate As Double: weightRate = CDbl(Nz(rsKeisu![加重率], 0))
        Dim kajuuPeriodNum As Long: kajuuPeriodNum = rsKeisu![荷重期数値]
        
        Dim sqlAct As String
        sqlAct = "SELECT [受注月], Sum([工事価格]) AS 合計額 " & _
                 "FROM [at_Icube_累計] " & _
                 "WHERE [一件工事判定] = '小口工事' " & _
                 "AND [施工管轄組織名] = '" & shopName & "' " & _
                 "AND [受注期] = " & kajuuPeriodNum & " " & _
                 "GROUP BY [受注月]"
        Set rsIn = db.OpenRecordset(sqlAct, dbOpenSnapshot)
        
        Do While Not rsIn.EOF
            Dim monthNum As Long: monthNum = rsIn![受注月]
            Dim weightedVal  As Currency: weightedVal = Int((Nz(rsIn![合計額], 0) * weightRate) + 0.5)
            
            Dim key As String: key = targetPeriod & "|" & shopName & "|" & monthNum
            
            If dictAgg.Exists(key) Then
                dictAgg(key) = dictAgg(key) + weightedVal
            Else
                dictAgg.Add key, weightedVal
            End If
            rsIn.MoveNext
        Loop
        rsIn.Close
        
        rsKeisu.MoveNext
    Loop
    rsKeisu.Close

    ' 3. テーブルへの書き出し
    Set rsOut = db.OpenRecordset(TARGET_TABLE, dbOpenDynaset)
    Dim k As Variant, parts As Variant
    For Each k In dictAgg.Keys
        parts = Split(k, "|")
        Dim mNum As Long: mNum = CLng(parts(2))
        
        rsOut.AddNew
        rsOut![予測ターゲット] = parts(0)
        rsOut![施工管轄組織名] = parts(1)
        
        ' Q判定 (4-6: 1Q, 7-9: 2Q, 10-12: 3Q, 1-3: 4Q)
        Select Case mNum
            Case 4, 5, 6: rsOut![Q] = "1Q"
            Case 7, 8, 9: rsOut![Q] = "2Q"
            Case 10, 11, 12: rsOut![Q] = "3Q"
            Case 1, 2, 3: rsOut![Q] = "4Q"
        End Select
        
        rsOut![受注月] = mNum & "月"
        rsOut![加重集計値] = dictAgg(k)
        rsOut.Update
    Next
    rsOut.Close
End Sub
