Attribute VB_Name = "acc_mod_KoguchiFinalForecast"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : acc_mod_KoguchiFinalForecast
' 概要           : 予測受注額と実績完工割合を統合し、最終的な予測完工高を算出するプログラム
' 更新日         : 2026/04/10 (仮コード自動付与対応)
'===================================================================================================

Private Const TARGET_TABLE  As String = AT_WORK_05_COMP_FCST
Private Const SRC_FORECAST  As String = AT_WORK_04_ORDER_FCST
Private Const SRC_PATTERN   As String = AT_WORK_03_COMP_RATIO

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Run_Final_Forecast_Reset
' 概要           : 加重平均予測受注額に完工割合を適用し、最終的な完工高予測テーブルを作成します。
'---------------------------------------------------------------------------------------------------
Public Sub Run_Final_Forecast_Reset(Optional isBatch As Boolean = False)
    On Error GoTo Err_Handler
    Dim db As DAO.Database: Set db = CurrentDb
    
    Debug.Print "--- 最終完工高予測 算出開始: " & Now & " ---"
    
    ' 1. テーブルの初期化
    Call Initialize_Target_Table(db)
    
    ' 2. データの集計と書き出し
    Call Aggregate_Final_Forecast_To_Table(db)
    
    Debug.Print "--- すべての工程が正常に完了しました ---"
    If Not isBatch Then
        MsgBox "最終予測テーブル [" & TARGET_TABLE & "] の作成が完了しました。", vbInformation, "完了"
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
          "期_予測ターゲット TEXT(20), " & _
          "施工管轄組織名 TEXT(100), " & _
          "受注月 TEXT(20), " & _
          "完工月 TEXT(20), " & _
          "完工Q TEXT(10), " & _
          "工事コード TEXT(50), " & _
          "元_受注予測額 CURRENCY, " & _
          "適用比率 DOUBLE, " & _
          "予測完工高 CURRENCY, " & _
          "予測経費額 CURRENCY)"
    db.Execute sql
End Sub

'---------------------------------------------------------------------------------------------------
' 2. 完工高予測の集計
'---------------------------------------------------------------------------------------------------
Private Sub Aggregate_Final_Forecast_To_Table(ByRef db As DAO.Database)
    Dim rsFor As DAO.Recordset: Dim rsPat As DAO.Recordset: Dim rsOut As DAO.Recordset
    Dim dictPatterns As Object: Set dictPatterns = CreateObject("Scripting.Dictionary")
    Dim dictOrg As Object: Set dictOrg = acc_mod_MappingTemplate.Get_Org_Dict()
    
    ' パターン割合をロード
    Set rsPat = db.OpenRecordset("SELECT [受注月], [完工月], [完工高割合] FROM [" & SRC_PATTERN & "]", dbOpenSnapshot)
    Do While Not rsPat.EOF
        Dim om As String: om = Nz(rsPat![受注月])
        Dim cm As String: cm = Nz(rsPat![完工月])
        Dim rt As Double: rt = CDbl(Nz(rsPat![完工高割合], 0))
        If Not dictPatterns.Exists(om) Then dictPatterns.Add om, CreateObject("Scripting.Dictionary")
        dictPatterns(om).Add cm, rt
        rsPat.MoveNext
    Loop
    rsPat.Close

    ' 共通経費率（14期固定）の取得
    Dim totalExpRate As Double: totalExpRate = 0
    Dim rsExp As DAO.Recordset
    Set rsExp = db.OpenRecordset("SELECT Sum(経費率) AS TotalRate FROM at_expBase WHERE 期 = '14期'", dbOpenSnapshot)
    If Not rsExp.EOF Then totalExpRate = Nz(rsExp!TotalRate, 0)
    rsExp.Close

    ' 予測受注額ループ
    Set rsFor = db.OpenRecordset("SELECT * FROM [" & SRC_FORECAST & "]", dbOpenSnapshot)
    Set rsOut = db.OpenRecordset(TARGET_TABLE, dbOpenDynaset)
    
    Do While Not rsFor.EOF
        Dim orderMonth As String: orderMonth = Nz(rsFor![受注月])
        Dim weightedAmt  As Currency: weightedAmt = Nz(rsFor![加重集計値], 0)
        Dim orgName      As String: orgName = Nz(rsFor![施工管轄組織名], "")
        
        If dictPatterns.Exists(orderMonth) Then
            Dim dictCM As Object: Set dictCM = dictPatterns(orderMonth)
            Dim finishMonth As Variant
            
            For Each finishMonth In dictCM.Keys
                Dim ratio As Double: ratio = dictCM(finishMonth)
                Dim mNumFinish As Integer: mNumFinish = val(finishMonth)
                Dim mNumOrder  As Integer: mNumOrder = val(orderMonth)
                
                ' 同期内順序チェック
                If ((mNumFinish + 8) Mod 12) < ((mNumOrder + 8) Mod 12) Then GoTo Skip_CM
                
                Dim finishQ As String
                Select Case mNumFinish
                    Case 4, 5, 6:    finishQ = "1Q"
                    Case 7, 8, 9:    finishQ = "2Q"
                    Case 10, 11, 12: finishQ = "3Q"
                    Case 1, 2, 3:    finishQ = "4Q"
                End Select
                
                ' --- 仮コードの取得 ---
                Dim projCode As String
                projCode = Get_TempProjectCode_Smart(db, dictOrg, orgName, finishQ)
                
                ' 金額計算
                Dim pAmt As Currency: pAmt = Int((weightedAmt * ratio) + 0.5)
                Dim pExp As Currency: pExp = Int((pAmt * totalExpRate) + 0.5)
                
                rsOut.AddNew
                rsOut![期_予測ターゲット] = rsFor![予測ターゲット]
                rsOut![施工管轄組織名] = orgName
                rsOut![受注月] = orderMonth
                rsOut![完工月] = finishMonth
                rsOut![完工Q] = finishQ
                rsOut![工事コード] = projCode
                rsOut![元_受注予測額] = weightedAmt
                rsOut![適用比率] = ratio
                rsOut![予測完工高] = pAmt
                rsOut![予測経費額] = pExp
                rsOut.Update
Skip_CM:
            Next
        End If
        rsFor.MoveNext
    Loop
    rsFor.Close: rsOut.Close
End Sub

'---------------------------------------------------------------------------------------------------
' 補助関数 : 組織名とQから仮工事コードを逆引きする
'---------------------------------------------------------------------------------------------------
Private Function Get_TempProjectCode_Smart(ByRef db As DAO.Database, ByRef dictOrg As Object, ByVal orgName As String, ByVal qStr As String) As String
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim qZen As String: qZen = StrConv(qStr, vbWide)

    ' -------------------------------------------------------------------------
    ' アプローチ1: フィールド [施工管轄組織名] を直接使った確実なマッチング (優先)
    ' -------------------------------------------------------------------------
    sql = "SELECT [仮基本工事コード] FROM [at_仮基本工事] " & _
          "WHERE [施工管轄組織名] = '" & orgName & "' " & _
          "AND ([基本工事名_Q] = '" & qZen & "' OR [基本工事名_Q] = '" & qStr & "' OR [基本工事名_Q] LIKE '*" & Left(qStr, 1) & "*') " & _
          "AND (Nz([基本工事名_官民], '') <> '官庁') " & _
          "AND (Nz([基本工事名_繰越], '') NOT LIKE '*繰越*') " & _
          "ORDER BY IIf([基本工事名_官民]='民間',0,1), Len([仮基本工事コード]) ASC"
          
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        Get_TempProjectCode_Smart = Nz(rs![仮基本工事コード], "99-FC-" & orgName)
        rs.Close
        Exit Function
    End If
    rs.Close

    ' -------------------------------------------------------------------------
    ' アプローチ2: マスタ(管轄作業所_RN部恒久作業所3) の [作業所_県] を使ったマッチング
    ' -------------------------------------------------------------------------
    Dim wsShort As String: wsShort = ""
    Dim rsState As DAO.Recordset
    
    ' 「盛岡」が「岩手」として登録されているケースに対応するため、県名フィールドを取得
    On Error Resume Next
    Set rsState = db.OpenRecordset("SELECT [作業所_県] FROM [at_管轄作業所_RN部恒久作業所3] WHERE [施工管轄組織名] = '" & orgName & "'", dbOpenSnapshot)
    If Err.Number = 0 Then
        If Not rsState.EOF Then
            wsShort = Nz(rsState![作業所_県], "")
        End If
        rsState.Close
    End If
    On Error GoTo 0
    
    ' 県名フィールドが空などの場合、旧ロジック(辞書逆引き)へフォールバック
    If wsShort = "" Then
        Dim key As Variant
        For Each key In dictOrg.Keys
            If dictOrg(key) = orgName Then
                wsShort = key: Exit For
            End If
        Next
        
        If wsShort = "" Then
            If InStr(orgName, "建築部") > 0 Then wsShort = "建築部"
        End If
        
        If wsShort = "" Then
            Get_TempProjectCode_Smart = "99-FC-" & orgName: Exit Function
        End If
        
        ' 「RN」を除去
        If UCase(Right(wsShort, 2)) = "RN" Then
            wsShort = Left(wsShort, Len(wsShort) - 2)
        ElseIf Right(wsShort, 2) = "ＲＮ" Then
            wsShort = Left(wsShort, Len(wsShort) - 2)
        End If
    End If
    
    ' 略称を使った検索
    sql = "SELECT [仮基本工事コード] FROM [at_仮基本工事] " & _
          "WHERE [基本工事名_作業所] = '" & wsShort & "' " & _
          "AND ([基本工事名_Q] = '" & qZen & "' OR [基本工事名_Q] = '" & qStr & "' OR [基本工事名_Q] LIKE '*" & Left(qStr, 1) & "*') " & _
          "AND (Nz([基本工事名_官民], '') <> '官庁') " & _
          "AND (Nz([基本工事名_繰越], '') NOT LIKE '*繰越*') " & _
          "ORDER BY IIf([基本工事名_官民]='民間',0,1), Len([仮基本工事コード]) ASC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        Get_TempProjectCode_Smart = Nz(rs![仮基本工事コード], "99-FC-" & wsShort)
    Else
        Get_TempProjectCode_Smart = "99-FC-" & wsShort
    End If
    rs.Close
End Function
