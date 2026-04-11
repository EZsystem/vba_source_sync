Attribute VB_Name = "acc_mod_KoguchiAggregationFinalizer"
'----------------------------------------------------------------
' Module: acc_mod_KoguchiAggregationFinalizer
' 説明   : 給与、実績経費、予測経費を統合し、判定ロジックを適用した最終集計テーブルを作成する
' 更新日 : 2026/04/10 (命名・プロセス統合対応)
'----------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const TARGET_TABLE As String = AT_WORK_FINAL_AGGREGATION

'----------------------------------------------------------------
' プロシージャ名 : Run_KoguchiStaffExpense_Aggregation
' 概要           : 最終集計プロセスを実行します
'----------------------------------------------------------------
Public Sub Run_KoguchiStaffExpense_Aggregation(Optional isBatch As Boolean = False)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim baseDate As Date
    
    On Error GoTo Err_Handler
    
    Debug.Print "--- 最終集計プロセス開始: " & Now & " ---"
    
    ' 1. 集計テーブルの初期化（クリアまたは再作成）
    Call Initialize_Aggregation_Table(db)
    
    ' 2. 最新実績日（受注計上日）の取得
    baseDate = Get_Latest_OrderDate(db)
    Debug.Print "    最新実績日（基準日）: " & Format(baseDate, "yyyy/mm/dd")
    
    ' 3. 実績データのインポート（給与）
    Call Insert_Actual_Salary(db)
    
    ' 4. 実績データのインポート（経費）
    Call Insert_Actual_Expenses(db)
    
    ' 5. 予測データのインポート（フィルタ適用）
    Call Insert_Forecast_Expenses(db, baseDate)
    
    Debug.Print "--- 集計プロセス正常終了 ---"
    
    If Not isBatch Then
        MsgBox "給与・経費の最終集計が完了しました。" & vbCrLf & _
               "集計基準日: " & Format(baseDate, "yyyy/mm/dd"), vbInformation
    End If
    Exit Sub

Err_Handler:
    MsgBox "集計エラー（" & Err.Number & "）: " & Err.Description, vbCritical
End Sub

'----------------------------------------------------------------
' 1. テーブル初期化：構造変更があるため一度削除して作り直す
'----------------------------------------------------------------
Private Sub Initialize_Aggregation_Table(ByRef db As DAO.Database)
    On Error Resume Next
    db.Execute "DROP TABLE [" & TARGET_TABLE & "]", dbFailOnError
    Err.Clear
    On Error GoTo 0
    
    Debug.Print "    テーブルを新規作成します: " & TARGET_TABLE
    Dim sql As String
    sql = "CREATE TABLE [" & TARGET_TABLE & "] (" & _
          "ID AUTOINCREMENT PRIMARY KEY, " & _
          "データ区分 TEXT(20), " & _
          "対象年月 DATETIME, " & _
          "作業所名 TEXT(100), " & _
          "工事コード TEXT(50), " & _
          "工事名 TEXT(255), " & _
          "社員名 TEXT(100), " & _
          "金額 CURRENCY)"
    db.Execute sql, dbFailOnError
End Sub

'----------------------------------------------------------------
' 2. 最新実績日の取得
'----------------------------------------------------------------
Private Function Get_Latest_OrderDate(ByRef db As DAO.Database) As Date
    Dim rs As DAO.Recordset
    ' at_Icube_累計 から最新の計上日を取得
    Set rs = db.OpenRecordset("SELECT Max(受注計上日_日付型) FROM at_Icube_累計", dbOpenSnapshot)
    If Not rs.EOF Then
        Get_Latest_OrderDate = Nz(rs.Fields(0), #1/1/2000#)
    Else
        Get_Latest_OrderDate = #1/1/2000#
    End If
    rs.Close
End Function

'----------------------------------------------------------------
' 3. 実績データのインポート（給与）
'----------------------------------------------------------------
Private Sub Insert_Actual_Salary(ByRef db As DAO.Database)
    Dim sql As String
    Debug.Print "    給与実績をインサート中..."
    sql = "INSERT INTO [" & TARGET_TABLE & "] (データ区分, 対象年月, 作業所名, 工事コード, 工事名, 社員名, 金額) " & _
          "SELECT '給与', K.年月, K.作業所名, K.工事コード, K.工事名, K.社員名, (G.本年度 * K.兼務率割合) AS 金額 " & _
          "FROM (at_kenmu AS K INNER JOIN at_社員情報 AS S ON K.社員名 = S.氏名_ﾒｰﾙ表示用) " & _
          "INNER JOIN _at_社員給与 AS G ON S.資格_想定給与額 = G.資格名 " & _
          "WHERE (((K.工事コード) <> '-') AND ((S.在籍区分) = True))"
    db.Execute sql
End Sub

'----------------------------------------------------------------
' 4. 実績データのインポート（経費）
'----------------------------------------------------------------
Private Sub Insert_Actual_Expenses(ByRef db As DAO.Database)
    Dim sql As String
    Debug.Print "    実績経費をインサート中..."
    ' at_工事経費_累計 には [工事コード] という列がないため [仮基本工事コード] を使用
    sql = "INSERT INTO [" & TARGET_TABLE & "] (データ区分, 対象年月, 作業所名, 工事コード, 工事名, 社員名, 金額) " & _
          "SELECT '実績経費', 年月, 作業所名, 仮基本工事コード, 工事名, 経費名, 経費額 " & _
          "FROM at_工事経費_累計 WHERE (仮基本工事コード <> 'EE0000100')"
    db.Execute sql
End Sub

'----------------------------------------------------------------
' 5. 予測データのインポート
'----------------------------------------------------------------
Private Sub Insert_Forecast_Expenses(ByRef db As DAO.Database, ByVal baseDate As Date)
    Dim rsFor As DAO.Recordset
    Dim rsOut As DAO.Recordset
    
    Debug.Print "    予測経費をフィルタリングしてインサート中..."
    
    Set rsFor = db.OpenRecordset("SELECT * FROM at_Work_05_完工_今期予測", dbOpenSnapshot)
    Set rsOut = db.OpenRecordset(TARGET_TABLE, dbOpenDynaset)
    
    Do While Not rsFor.EOF
        Dim termStr As String: termStr = Nz(rsFor![期_予測ターゲット], "")
        Dim monthStr As String: monthStr = Nz(rsFor![完工月], "")
        
        ' 期と月から正確な日付（月初）を算出
        Dim finishDate As Date: finishDate = Convert_TermMonth_To_Date(termStr, monthStr)
        
        ' --- 判定ロジック: 基準日より「後」の月のみ算入 ---
        If finishDate > baseDate Then
            rsOut.AddNew
            rsOut![データ区分] = "予測経費"
            rsOut![対象年月] = finishDate
            
            ' 予測用の工事コードを取得
            Dim projCode As String: projCode = Nz(rsFor![工事コード], "99-FORECAST")
            Dim finalOrgName As String: finalOrgName = Nz(rsFor![施工管轄組織名], "")
            Dim finalProjName As String: finalProjName = "【予測】小口完工"
            
            ' --- 工事コードを使ってマスタから正確な作業所名と工事名を取得 ---
            If projCode <> "99-FORECAST" Then
                Dim rsMaster As DAO.Recordset
                Set rsMaster = db.OpenRecordset("SELECT 施工管轄組織名, 仮基本工事名称 FROM at_仮基本工事 WHERE 仮基本工事コード = '" & projCode & "'", dbOpenSnapshot)
                If Not rsMaster.EOF Then
                    finalOrgName = Nz(rsMaster![施工管轄組織名], finalOrgName)
                    finalProjName = Nz(rsMaster![仮基本工事名称], finalProjName)
                    
                    ' --- ？？ を年度（全角2桁）に置換 ---
                    If InStr(finalProjName, "？？") > 0 And termStr <> "" Then
                        Dim termNum As Integer: termNum = Val(Replace(termStr, "期", ""))
                        If termNum > 0 Then
                            ' 14期 -> 2026年度 -> 26 -> ２６
                            Dim yyStr As String: yyStr = StrConv(Format(termNum + 12, "00"), vbWide)
                            finalProjName = Replace(finalProjName, "？？", yyStr)
                        End If
                    End If
                End If
                rsMaster.Close
            End If
            
            rsOut![作業所名] = finalOrgName
            rsOut![工事コード] = projCode
            rsOut![工事名] = finalProjName
            rsOut![社員名] = "予測経費"
            rsOut![金額] = Nz(rsFor![予測経費額], 0)
            rsOut.Update
        End If
        
        rsFor.MoveNext
    Loop
    
    rsFor.Close: rsOut.Close
End Sub

'----------------------------------------------------------------
' 補助関数 : 期(XX期)と月(X月)から日付(西暦月初)に変換
'----------------------------------------------------------------
Private Function Convert_TermMonth_To_Date(ByVal termStr As String, ByVal monthStr As String) As Date
    Dim termNum As Integer
    Dim mNum As Integer
    Dim yNum As Integer
    
    termNum = Val(Replace(termStr, "期", ""))
    mNum = Val(Replace(monthStr, "月", ""))
    
    If termNum = 0 Or mNum = 0 Then Exit Function
    
    yNum = termNum + 2012
    If mNum <= 3 Then
        yNum = yNum + 1
    End If
    
    Convert_TermMonth_To_Date = DateSerial(yNum, mNum, 1)
End Function


