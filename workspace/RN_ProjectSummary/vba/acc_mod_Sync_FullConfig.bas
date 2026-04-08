Attribute VB_Name = "acc_mod_Sync_FullConfig"
'----------------------------------------------------------------
' Module: acc_mod_Sync_FullConfig
' 概要   : _at_ExportConfig テーブルの全レコードを最新のクエリ名・SQLテンプレートに更新する
'----------------------------------------------------------------
Option Compare Database
Option Explicit

Public Sub Run_Registry_Sync_Full()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim count As Long
    On Error GoTo Err_Handler

    ' --- シート: IcubeData ---
    Call Upsert(db, 2, "IcubeData", "q_受注_月計_小口", "xl_IcubeJyu", count, _
        "SELECT First(at_Icube_累計.基本工事コード) AS 基本工事コード, First(at_Icube_累計.基本工事名称) AS 基本工事名称, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計, at_Icube_累計.施工管轄組織名, First([受注期] & '期') AS 受注期表示, First([受注Q] & 'Q') AS 受注Q表示, [受注月] & '月' AS 受注月表示, at_Icube_累計.基本工事名_官民 FROM at_Icube_累計 WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND ([受注期]=Val(Replace('{TERM}','期',''))) AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事')) GROUP BY at_Icube_累計.施工管轄組織名, [受注月] & '月', at_Icube_累計.基本工事名_官民;")
    
    Call Upsert(db, 4, "IcubeData", "q_受注_月計_一件", "xl_IcubeIken", count, _
        "SELECT [at_Icube_累計].工事コード, First([at_Icube_累計].工事名称) AS 工事名称, Sum([at_Icube_累計].工事価格) AS 工事価格_合計, Sum([at_Icube_累計].粗利益額) AS 粗利益額_合計, First([at_Icube_累計].施工管轄組織名) AS 施工管轄組織名, First([受注期] & '期') AS 受注期表示, First([受注Q] & 'Q') AS 受注Q表示, [受注月] & '月' AS 受注月表示 FROM at_Icube_累計 WHERE ((([at_Icube_累計].施工管轄組織名)<>'ビルサービスグループ') And (([受注期] & '期')='{TERM}') And (([at_Icube_累計].一件工事判定)='一件工事') And (([at_Icube_累計].所属組織名)='ＬＣＳ事業部')) GROUP BY [at_Icube_累計].工事コード, [受注月] & '月';")
    
    Call Upsert(db, 10, "IcubeData", "q_受注_Q計_小口", "xl_IcubeJyuSum", count, _
        "SELECT at_Icube_累計.施工管轄組織名, ([受注期] & '期') AS 受注期表示, ([受注Q] & 'Q') AS 受注Q表示, at_Icube_累計.基本工事名_官民, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計 FROM at_Icube_累計 WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND ([受注期]=Val(Replace('{TERM}','期',''))) AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事')) GROUP BY at_Icube_累計.施工管轄組織名, ([受注期] & '期'), ([受注Q] & 'Q'), at_Icube_累計.基本工事名_官民;")
    
    Call Upsert(db, 11, "IcubeData", "q_受注_月計_形態別_小口", "xl_IcubeJyuKeitaSum", count)
    Call Upsert(db, 12, "IcubeData", "q_受注_月計_区分なし", "xl_IcubeJyuMon2Sum", count)
    Call Upsert(db, 13, "IcubeData", "q_小口受注完工推移_3期分", "xl_IcubeJyu3year", count)
    Call Upsert(db, 14, "IcubeData", "q_受注完工予測_加重平均集計", "xl_IcubeJyu3yearKajyu", count)

    ' --- シート: 経費M ---
    Call Upsert(db, 15, "経費M", "sqlSum_給与経費月毎", "xt_expSumMon", count)
    Call Upsert(db, 16, "経費M", "sqlSum_兼務率職員単位月毎", "xt_expEmpSumMon", count)
    Call Upsert(db, 6, "経費M", "sel_expFront", "xt_expFront", count, "SELECT * FROM at_expFront WHERE (期)='{TERM}';")
    Call Upsert(db, 7, "経費M", "sel_expBase", "xt_expBase", count, "SELECT * FROM at_expBase WHERE (期)='{TERM}';")
    Call Upsert(db, 8, "経費M", "sel_expDraftRate", "xt_expDraftRate", count, "SELECT * FROM at_expDraftRate WHERE (期)='{TERM}';")

    ' --- シート: ac_受注完工予測 ---
    Call Upsert(db, 17, "ac_受注完工予測", "at_Work_受注完工予測_加重平均集計", "xt_JyuYosoku", count)
    Call Upsert(db, 18, "ac_受注完工予測", "at_Work_予測完工高_最終結果", "xt_KanYosoku", count)
    Call Upsert(db, 19, "ac_受注完工予測", "q_繰越工事集計", "xt_JyuKuri", count, "SELECT * FROM q_繰越工事集計;")
    Call Upsert(db, 20, "ac_受注完工予測", "(未作成)", "xt_expSumKisyu", count)
    Call Upsert(db, 21, "ac_受注完工予測", "sqlSum_兼務率工事単位", "xt_expSumKisyuIken", count)

    MsgBox count & " 件の設定を同期しました。", vbInformation
    Exit Sub
Err_Handler:
    MsgBox "更新エラー: " & Err.Description, vbCritical
End Sub

Private Sub Upsert(db As DAO.Database, id As Long, sn As String, qry As String, tbl As String, ByRef c As Long, Optional sqlTmp As String = "")
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM [_at_ExportConfig] WHERE [ID] = " & id)
    If rs.EOF Then
        rs.AddNew: rs!ID = id
    Else
        rs.Edit
    End If
    rs!ProcessName = "RN管理表エクスポート"
    rs!ExcelSheet = sn
    rs!queryName = qry
    rs!ExcelTable = tbl
    rs!sqlTemplate = sqlTmp
    rs!IsActive = True
    rs.Update
    rs.Close: c = c + 1
End Sub
