Attribute VB_Name = "acc_mod_Sync_FullConfig"
'----------------------------------------------------------------
' 概要 : _at_ExportConfig テーブルの全21レコードを完全同期
' 更新 : 予測テーブルの LIKE '{TERM}*' 絞り込みに対応
'----------------------------------------------------------------
Public Sub Final_Registry_Sync_v3_2()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim p As String: p = "D:\My_code\11_workspaces\RN_kanri_system\RNkanri_system\ＲＮ管理表14期期首.xlsm"
    Dim pn As String: pn = "RN管理表エクスポート"
    
    On Error GoTo Err_Handler
    db.Execute "DELETE FROM [_at_ExportConfig]", dbFailOnError
    
    ' --- 1. 原価基本 ---
    Call FullIns(db, 1, pn, "sel_原価S_基本工事", "SELECT kt.仮基本工事コード, kt.仮基本工事略名, g.基本工事コード, g.基本工事名, g.工事価格, g.[工事原価(経費込)], g.予定利益, g.粗利率, g.直接工事費, g.経費, g.作業所経費, g.率, g.共通経費, g.率2, g.[既払高：総額], g.[既払高：経費], g.今後支払予定, g.当月より前の支払金額, g.当月支払金額, g.[設計料・他], g.当月以降予定金額, g.行分類 FROM at_原価S_基本工事 AS g INNER JOIN at_Icube_累計 AS kt ON g.基本工事コード = kt.s基本工事コード GROUP BY kt.仮基本工事コード, kt.仮基本工事略名, g.基本工事コード, g.基本工事名, g.工事価格, g.[工事原価(経費込)], g.予定利益, g.粗利率, g.直接工事費, g.経費, g.作業所経費, g.率, g.共通経費, g.率2, g.[既払高：総額], g.[既払高：経費], g.今後支払予定, g.当月より前の支払金額, g.当月支払金額, g.[設計料・他], g.当月以降予定金額, g.行分類;", p, "原価Data", "xl_genkaKihon")

    ' --- 2. 受注月計 (全セット) ---
    Call FullIns(db, 2, pn, "q_受注_月計_小口", "SELECT First(at_Icube_累計.基本工事コード) AS 基本工事コード, First(at_Icube_累計.基本工事名称) AS 基本工事名称, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計, at_Icube_累計.施工管轄組織名, First([受注期] & '期') AS 受注期表示, First([受注Q] & 'Q') AS 受注Q表示, [受注月] & '月' AS 受注月表示, at_Icube_累計.基本工事名_官民 FROM at_Icube_累計 WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND ([受注期]=Val(Replace('{TERM}','期',''))) AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事')) GROUP BY at_Icube_累計.施工管轄組織名, [受注月] & '月', at_Icube_累計.基本工事名_官民;", p, "IcubeData", "xl_IcubeJyu")
    Call FullIns(db, 3, pn, "q_完工_月計_小口", "SELECT at_Icube_累計.基本工事コード, First(at_Icube_累計.基本工事名称) AS 基本工事名称, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計, First(at_Icube_累計.施工管轄組織名) AS 施工管轄組織名, First([完工期] & '期') AS 完工期表示, First([完工Q] & 'Q') AS 完工Q表示, [完工月] & '月' AS 完工月表示, at_Icube_累計.基本工事名_官民 FROM at_Icube_累計 WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND (([完工期] & '期')='{TERM}') AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事')) GROUP BY at_Icube_累計.基本工事コード, [完工月] & '月', at_Icube_累計.基本工事名_官民;", p, "IcubeData", "xl_IcubeKan")
    Call FullIns(db, 4, pn, "q_受注_月計_一件", "SELECT [at_Icube_累計].工事コード, First([at_Icube_累計].工事名称) AS 工事名称, Sum([at_Icube_累計].工事価格) AS 工事価格_合計, Sum([at_Icube_累計].粗利益額) AS 粗利益額_合計, First([at_Icube_累計].施工管轄組織名) AS 施工管轄組織名, First([受注期] & '期') AS 受注期表示, First([受注Q] & 'Q') AS 受注Q表示, [受注月] & '月' AS 受注月表示 FROM at_Icube_累計 WHERE ((([at_Icube_累計].施工管轄組織名)<>'ビルサービスグループ') And (([受注期] & '期')='{TERM}') And (([at_Icube_累計].一件工事判定)='一件工事') And (([at_Icube_累計].所属組織名)='ＬＣＳ事業部')) GROUP BY [at_Icube_累計].工事コード, [受注月] & '月';", p, "IcubeData", "xl_IcubeIken")
    Call FullIns(db, 5, pn, "q_受注_月計_建築", "SELECT First([at_Icube_累計].施工管轄組織名) AS 施工管轄組織名, First([受注期] & '期') AS 受注期表示, [受注Q] & 'Q' AS 受注Q表示, Sum([at_Icube_累計].工事価格) AS 工事価格_合計, Sum([at_Icube_累計].粗利益額) AS 粗利益額_合計, First([at_Icube_累計].一件工事判定) AS 一件工事判定, First([at_Icube_累計].[リニューアル環境区分名]) AS リニューアル環境区分名 FROM at_Icube_累計 WHERE ((([受注期] & '期')='{TERM}') And (([at_Icube_累計].[リニューアル環境区分名])='リニューアル') And (([at_Icube_累計].所属組織名)='建築部')) GROUP BY [受注Q] & 'Q';", p, "IcubeData", "xl_IcubeKent")

    ' --- 3. 経費率・管理シート ---
    Call FullIns(db, 6, pn, "sel_expFront", "SELECT * FROM at_expFront WHERE (期)='{TERM}';", p, "経費M", "xt_expFront")
    Call FullIns(db, 7, pn, "sel_expBase", "SELECT * FROM at_expBase WHERE (期)='{TERM}';", p, "経費M", "xt_expBase")
    Call FullIns(db, 8, pn, "sel_expDraftRate", "SELECT * FROM at_expDraftRate WHERE (期)='{TERM}';", p, "経費M", "xt_expDraftRate")

    ' --- 4. Q計・新規項目 (IcubeData) ---
    Call FullIns(db, 9, pn, "q_完工_Q計_小口", "SELECT at_Icube_累計.施工管轄組織名, ([完工期] & '期') AS 完工期表示, ([完工Q] & 'Q') AS 完工Q表示, at_Icube_累計.基本工事名_官民, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計 FROM at_Icube_累計 WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND (([完工期] & '期')='{TERM}') AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事')) GROUP BY at_Icube_累計.施工管轄組織名, ([完工期] & '期'), ([完工Q] & 'Q'), at_Icube_累計.基本工事名_官民;", p, "IcubeData", "xl_IcubeKanSum")
    Call FullIns(db, 10, pn, "q_受注_Q計_小口", "SELECT at_Icube_累計.施工管轄組織名, ([受注期] & '期') AS 受注期表示, ([受注Q] & 'Q') AS 受注Q表示, at_Icube_累計.基本工事名_官民, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計 FROM at_Icube_累計 WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND ([受注期]=Val(Replace('{TERM}','期',''))) AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事')) GROUP BY at_Icube_累計.施工管轄組織名, ([受注期] & '期'), ([受注Q] & 'Q'), at_Icube_累計.基本工事名_官民;", p, "IcubeData", "xl_IcubeJyuSum")
    
    Call FullIns(db, 11, pn, "q_受注_月計_形態別_小口", "SELECT at_Icube_累計.施工管轄組織名, ([受注期] & '期') AS 受注期表示, ([受注Q] & 'Q') AS 受注Q表示, [受注月] & '月' AS 受注月表示, at_Icube_累計.受注形態名, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計 FROM at_Icube_累計 WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND ([受注期]=Val(Replace('{TERM}','期',''))) AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事')) GROUP BY at_Icube_累計.施工管轄組織名, ([受注期] & '期'), ([受注Q] & 'Q'), [受注月] & '月', at_Icube_累計.受注形態名;", p, "IcubeData", "xl_IcubeJyuKeitaSum")
    Call FullIns(db, 12, pn, "q_受注_月計_区分なし", "SELECT at_Icube_累計.施工管轄組織名, ([受注期] & '期') AS 受注期表示, ([受注Q] & 'Q') AS 受注Q表示, [受注月] & '月' AS 受注月表示, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計 FROM at_Icube_累計 WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND ([受注期]=Val(Replace('{TERM}','期',''))) AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事')) GROUP BY at_Icube_累計.施工管轄組織名, ([受注期] & '期'), ([受注Q] & 'Q'), [受注月] & '月';", p, "IcubeData", "xl_IcubeJyuMon2Sum")
    Call FullIns(db, 13, pn, "q_小口受注完工推移_3期分", "SELECT [施工管轄組織名], [受注期] & '期' AS 期, Sum([工事価格]) AS 合計工事価格, Sum([粗利益額]) AS 合計粗利益額 FROM at_Icube_累計 WHERE ([一件工事判定] = '小口工事') AND ([所属組織名] = 'ＬＣＳ事業部') GROUP BY [施工管轄組織名], [受注期];", p, "IcubeData", "xl_IcubeJyu3year")
    Call FullIns(db, 14, pn, "q_04_受注_今期予測", "SELECT * FROM at_Work_04_受注_今期予測;", p, "IcubeData", "xl_IcubeJyu3yearKajyu")
    
    ' --- 5. 経費・予測シート (新規) ---
    Call FullIns(db, 15, pn, "sqlSum_給与経費月毎", "", p, "経費M", "xt_expSumMon")
    Call FullIns(db, 16, pn, "sqlSum_兼務率職員単位月毎", "", p, "経費M", "xt_expEmpSumMon")

    ' --- ★予測テーブルの絞り込みエクスポート設定 ---
    Call FullIns(db, 17, pn, "at_Work_04_受注_今期予測", "SELECT * FROM [at_Work_04_受注_今期予測] WHERE [予測ターゲット] LIKE '{TERM}*';", p, "ac_受注完工予測", "xt_JyuYosoku")
    Call FullIns(db, 18, pn, "at_Work_05_完工_今期予測", "SELECT * FROM [at_Work_05_完工_今期予測] WHERE [期_予測ターゲット] LIKE '{TERM}*';", p, "ac_受注完工予測", "xt_KanYosoku")
    
    Call FullIns(db, 19, pn, "q_繰越工事集計", "SELECT [at_Icube_累計].[施工管轄組織名], [at_Icube_累計].[完工期] & '期' AS 期, Sum([at_Icube_累計].[工事価格]) AS 合計工事価格, Sum([at_Icube_累計].[粗利益額]) AS 合計粗利益額 FROM at_Icube_累計 WHERE ([at_Icube_累計].[一件工事判定] = '小口工事') AND ([at_Icube_累計].[基本工事名_繰越] = '(繰越)') AND ([at_Icube_累計].[所属組織名] = 'ＬＣＳ事業部') AND ([at_Icube_累計].[完工期] = Val(Replace('{TERM}','期',''))) GROUP BY [at_Icube_累計].[施工管轄組織名], [at_Icube_累計].[完工期];", p, "ac_受注完工予測", "xt_JyuKuri")
    Call FullIns(db, 20, pn, "(未作成)", "", p, "ac_受注完工予測", "xt_expSumKisyu")
    Call FullIns(db, 21, pn, "sqlSum_兼務率工事単位", "", p, "ac_受注完工予測", "xt_expSumKisyuIken")

    MsgBox "全21件のRegistry完全同期（予測テーブル絞り込み対応版）が完了しました。", vbInformation
    Exit Sub
Err_Handler:
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Private Sub FullIns(db As DAO.Database, ID As Long, pn As String, qry As String, sql As String, path As String, sn As String, tbl As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM [_at_ExportConfig] WHERE 1=0", dbOpenDynaset)
    rs.AddNew
    rs!ID = ID: rs!ProcessName = pn: rs!queryName = qry: rs!sqlTemplate = sql: rs!ExcelPath = path: rs!ExcelSheet = sn: rs!ExcelTable = tbl: rs!IsActive = True
    rs.Update: rs.Close
End Sub

