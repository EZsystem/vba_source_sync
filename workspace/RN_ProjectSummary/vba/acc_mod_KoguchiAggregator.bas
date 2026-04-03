Attribute VB_Name = "acc_mod_KoguchiAggregator"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : acc_mod_KoguchiAggregator
' 概要           : 小口工事の受注比率(B/A)をワークテーブルを用いて段階的に集計・保存
'                : 1ステップごとに値をテーブルに保存し、透明性を確保します。
'===================================================================================================

' ワークテーブル名の定義
Private Const WT_TRANSITION As String = "at_Work_小口完工推移"
Private Const WT_WEIGHTED_SUM As String = "at_Work_加重平均集計"
Private Const WT_FINAL_RATIO As String = "at_Work_小口受注比率_最終"

'---------------------------------------------------------------------------------------------------
' ステップ1: 実績データのテーブル化 (q_小口受注完工推移_3期分 -> at_Work_小口完工推移)
'---------------------------------------------------------------------------------------------------
Public Sub Execute_Step1_Actuals()
    On Error GoTo Err_Handler
    Dim db As DAO.Database: Set db = CurrentDb
    
    ' テーブルの存在確認と作成
    Call Ensure_Single_Table_Exists(db, WT_TRANSITION)
    
    ' クリア
    db.Execute "DELETE FROM [" & WT_TRANSITION & "]", dbFailOnError
    
    ' 書き出し (INSERT INTO)
    Dim sql As String
    sql = "INSERT INTO [" & WT_TRANSITION & "] ( [施工管轄組織名], [受注期_], [完工期_], [受注月_], [完工月_], [工事価格の合計] ) " & _
          "SELECT [施工管轄組織名], [受注期_], [完工期_], [受注月_], [完工月_], [工事価格の合計] FROM [q_小口受注完工推移_3期分]"
    db.Execute sql, dbFailOnError
    
    MsgBox "ステップ1完了: 実績データの書き出しが完了しました。" & vbCrLf & _
           "テーブル [" & WT_TRANSITION & "] を開き、内容に不足がないかご確認ください。", vbInformation
    Exit Sub
Err_Handler:
    MsgBox "ステップ1エラー: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------------------------------------
' ステップ2: 予測ベースのテーブル化 (q_受注完工予測_加重平均集計 -> at_Work_加重平均集計)
'---------------------------------------------------------------------------------------------------
Public Sub Execute_Step2_Weighted()
    On Error GoTo Err_Handler
    Dim db As DAO.Database: Set db = CurrentDb
    
    ' テーブルの存在確認と作成
    Call Ensure_Single_Table_Exists(db, WT_WEIGHTED_SUM)
    
    ' クリア
    db.Execute "DELETE FROM [" & WT_WEIGHTED_SUM & "]", dbFailOnError
    
    ' 書き出し
    Dim sql As String
    sql = "INSERT INTO [" & WT_WEIGHTED_SUM & "] ( [期_計算対象], [期_計算対象数値], [施工管轄組織名], [Q], [受注月数値], [加重受注高] ) " & _
          "SELECT [期_計算対象], [期_計算対象数値], [施工管轄組織名], [Q], [受注月数値], [加重受注高] FROM [q_受注完工予測_加重平均集計]"
    db.Execute sql, dbFailOnError
    
    MsgBox "ステップ2完了: 予測ベースの書き出しが完了しました。" & vbCrLf & _
           "テーブル [" & WT_WEIGHTED_SUM & "] で数値を確認してください。", vbInformation
    Exit Sub
Err_Handler:
    MsgBox "ステップ2エラー: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------------------------------------
' ステップ3: 最終比率の計算と保存
'---------------------------------------------------------------------------------------------------
Public Sub Execute_Step3_Ratio()
    On Error GoTo Err_Handler
    Dim db As DAO.Database: Set db = CurrentDb
    
    ' テーブルの存在確認と作成
    Call Ensure_Single_Table_Exists(db, WT_FINAL_RATIO)
    
    ' クリア
    db.Execute "DELETE FROM [" & WT_FINAL_RATIO & "]", dbFailOnError
    
    ' 最終集計ロジック (全社合計)
    Dim sql As String
    sql = "INSERT INTO [" & WT_FINAL_RATIO & "] ( 計算日時, 対象期, 受注月, 実績A_実数値, 予測B_期待値, 受注比率 ) " & _
          "SELECT Now(), B.期, B.月, CDbl(Nz(A.実績A, 0)), CDbl(Nz(B.予測B, 0)), CDbl(Nz(B.予測B, 0) / IIf(Nz(A.実績A, 0) = 0, Null, A.実績A)) " & _
          "FROM (" & _
          "  SELECT 受注月数値 AS 月, Sum(加重受注高) AS 予測B, Max(期_計算対象数値) AS 期 " & _
          "  FROM [" & WT_WEIGHTED_SUM & "] " & _
          "  WHERE 期_計算対象数値 = (SELECT Max(Val(受注期_)) FROM [" & WT_TRANSITION & "]) " & _
          "  GROUP BY 受注月数値" & _
          ") AS B " & _
          "LEFT JOIN (" & _
          "  SELECT Val(受注月_) AS 月, Sum(工事価格の合計) AS 実績A " & _
          "  FROM [" & WT_TRANSITION & "] " & _
          "  WHERE (受注月_ = 完工月_) AND (Val(受注期_) = (SELECT Max(Val(受注期_)) FROM [" & WT_TRANSITION & "])) " & _
          "  GROUP BY Val(受注月_)" & _
          ") AS A ON B.月 = A.月"
    
    db.Execute sql, dbFailOnError
    
    MsgBox "全工程完了: 比率計算が完了しました。" & vbCrLf & _
           "テーブル [" & WT_FINAL_RATIO & "] をご確認ください。", vbInformation
    Exit Sub
Err_Handler:
    MsgBox "ステップ3エラー: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------------------------------------------
' 補助関数群 (テーブル作成)
'---------------------------------------------------------------------------------------------------
Private Sub Ensure_Single_Table_Exists(ByRef db As DAO.Database, ByVal tName As String)
    On Error Resume Next
    db.OpenRecordset("SELECT TOP 1 * FROM [" & tName & "]").Close
    If Err.Number <> 0 Then
        Err.Clear
        Dim sql As String
        Select Case tName
            Case WT_WEIGHTED_SUM: sql = "CREATE TABLE [" & tName & "] (ID AUTOINCREMENT PRIMARY KEY, 期_計算対象 TEXT(50), 期_計算対象数値 LONG, 施工管轄組織名 TEXT(255), Q TEXT(10), 受注月数値 LONG, 加重受注高 CURRENCY)"
            Case WT_TRANSITION:   sql = "CREATE TABLE [" & tName & "] (ID AUTOINCREMENT PRIMARY KEY, 施工管轄組織名 TEXT(255), 受注期_ TEXT(20), 完工期_ TEXT(20), 受注月_ TEXT(20), 完工月_ TEXT(20), 工事価格の合計 CURRENCY)"
            Case WT_FINAL_RATIO:  sql = "CREATE TABLE [" & tName & "] (ID AUTOINCREMENT PRIMARY KEY, 計算日時 DATETIME, 対象期 LONG, 受注月 LONG, 実績A_実数値 CURRENCY, 予測B_期待値 CURRENCY, 受注比率 DOUBLE)"
        End Select
        db.Execute sql
    End If
End Sub
