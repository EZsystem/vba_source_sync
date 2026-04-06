Attribute VB_Name = "acc_mod_KoguchiAggregator"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : acc_mod_KoguchiAggregator
' 概要           : 実績推移テーブル (5項目) を全社合計で作成・更新するプログラム
' 更新日         : 2026/04/03
'===================================================================================================

' Private Const TARGET_TABLE As String = "at_Work_01_実績推移_3期分"
Private Const TARGET_TABLE As String = AT_WORK_01_ACTUALS_3P

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Run_Aggregator_Reset
' 概要           : 実績推移テーブルを最新の3期分で再構築します。
'---------------------------------------------------------------------------------------------------
Public Sub Run_Aggregator_Reset(Optional isBatch As Boolean = False)
    On Error GoTo Err_Handler
    Dim db As DAO.Database: Set db = CurrentDb
    
    Debug.Print "--- 実績推移集計開始: " & Now & " ---"
    
    ' 1. テーブルの初期化 (削除して再作成)
    Call Initialize_Target_Table(db)
    
    ' 2. 実績データの集計
    Call Aggregate_Actuals_To_Table(db)
    
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
' 1. テーブル初期化 (物理的な再構築)
'---------------------------------------------------------------------------------------------------
Private Sub Initialize_Target_Table(ByRef db As DAO.Database)
    On Error Resume Next
    db.Execute "DROP TABLE [" & TARGET_TABLE & "]"
    Err.Clear
    On Error GoTo 0
    
    Debug.Print "テーブルを新規作成中: " & TARGET_TABLE
    ' ご提示いただいた5項目の構造
    Dim sql As String
    sql = "CREATE TABLE [" & TARGET_TABLE & "] (" & _
          "ID AUTOINCREMENT PRIMARY KEY, " & _
          "受注期_ TEXT(20), " & _
          "完工期_ TEXT(20), " & _
          "受注月_ TEXT(20), " & _
          "完工月_ TEXT(20), " & _
          "工事価格の合計 CURRENCY)"
    db.Execute sql
End Sub

'---------------------------------------------------------------------------------------------------
' 2. データの集計と書き込み
'---------------------------------------------------------------------------------------------------
Private Sub Aggregate_Actuals_To_Table(ByRef db As DAO.Database)
    Dim rsIn As DAO.Recordset
    Dim rsOut As DAO.Recordset
    Dim dictAgg As Object: Set dictAgg = CreateObject("Scripting.Dictionary")
    Dim targetPeriodMax As Long
    Dim strSQL As String
    
    ' 最新の期を特定
    targetPeriodMax = db.OpenRecordset("SELECT Max([受注期]) FROM [at_Icube_累計]").Fields(0)
    Debug.Print "最新期: " & targetPeriodMax & "期"
    
    ' 元データの抽出 (小口工事 かつ 推奨6作業所)
    strSQL = "SELECT * FROM [at_Icube_累計] " & _
             "WHERE [一件工事判定] = '小口工事' " & _
             "AND [施工管轄組織名] IN ('青森ＲＮ（作）', '秋田ＲＮ（作）', '盛岡ＲＮ（作）', '山形ＲＮ（作）', '福島ＲＮ（作）', '仙台ＲＮ（作）') " & _
             "AND [受注期] >= " & (targetPeriodMax - 2)
    
    Set rsIn = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    ' --- VBAメモリ集計ループ ---
    Dim key As String, val As Currency
    Do While Not rsIn.EOF
        key = Nz(rsIn![受注期]) & "|" & _
              Nz(rsIn![完工期]) & "|" & _
              Nz(rsIn![受注月]) & "|" & _
              Nz(rsIn![完工月])
        
        val = Nz(rsIn![工事価格], 0)
        
        If dictAgg.Exists(key) Then
            dictAgg(key) = dictAgg(key) + val
        Else
            dictAgg.Add key, val
        End If
        rsIn.MoveNext
    Loop
    
    ' --- テーブルへの書き出し ---
    Set rsOut = db.OpenRecordset(TARGET_TABLE, dbOpenDynaset)
    Dim k As Variant, parts As Variant
    For Each k In dictAgg.Keys
        parts = Split(k, "|")
        rsOut.AddNew
        rsOut![受注期_] = parts(0) & "期"
        rsOut![完工期_] = parts(1) & "期"
        rsOut![受注月_] = parts(2) & "月"
        rsOut![完工月_] = parts(3) & "月"
        rsOut![工事価格の合計] = dictAgg(k)
        rsOut.Update
    Next
    
    rsIn.Close: rsOut.Close
End Sub
