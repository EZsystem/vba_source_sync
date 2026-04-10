Attribute VB_Name = "acc_mod_ExcelExporter"
'----------------------------------------------------------------
' Module: acc_mod_ExcelExporter
' 説明   : 管理テーブル (_at_ExportConfig) 駆動型 Excel エクスポート (診断機能付)
' 更新日 : 2026/04/09 (完全復旧 & インテリジェント・マッピング版)
'----------------------------------------------------------------
Option Compare Database
Option Explicit

' --- 設定定数 ---
Private Const CLOSE_EXCEL_AFTER_EXPORT As Boolean = False

'----------------------------------------------------------------
' プロシージャ名 : Execute_Excel_Data_Export
'----------------------------------------------------------------
Public Sub Execute_Excel_Data_Export()
    Dim db      As DAO.Database: Set db = CurrentDb
    Dim rs      As DAO.Recordset
    Dim xlApp   As Object
    Dim xlBook  As Object
    Dim xlSheet As Object
    Dim targetTerm As String
    Dim openWBs As Object
    
    On Error GoTo Err_Handler
    
    ' 診断のため Echo は True にしておきます
    Application.Echo True
    
    Debug.Print "=== エクスポート診断開始: " & Now & " ==="
    
    ' 1. フォームから期を取得
    If CurrentProject.AllForms("frm_SystemMain").IsLoaded Then
        targetTerm = Nz(Forms("frm_SystemMain")!cmbTargetTerm.Value, "")
    End If
    
    If targetTerm = "" Then
        MsgBox "期が選択されていません。", vbExclamation, "診断エラー"
        Exit Sub
    End If
    Debug.Print "  ターゲット期: " & targetTerm
    
    ' 2. 管理テーブルから有効な設定を取得
    ' ※ acc_mod_MappingTemplate の AT_EXPORT_CONFIG を参照
    Set rs = db.OpenRecordset("SELECT * FROM [" & AT_EXPORT_CONFIG & "] WHERE [IsActive] = True ORDER BY [ID]", dbOpenSnapshot)
    If rs.EOF Then
        MsgBox "有効な設定(IsActive=True)がありません。", vbInformation
        rs.Close: Exit Sub
    End If
    
    ' 3. Excelアプリケーションの準備
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set openWBs = CreateObject("Scripting.Dictionary")
    
    ' 4. ループ実行
    Do While Not rs.EOF
        Dim currentID As Long: currentID = rs!ID
        Dim procName  As String: procName = Nz(rs!ProcessName, "")
        Dim qryName   As String: qryName = Nz(rs!queryName, "")
        Dim sqlTemplate As String: sqlTemplate = Nz(rs!sqlTemplate, "")
        Dim xlPath    As String: xlPath = Nz(rs!ExcelPath, "")
        Dim snName    As String: snName = Nz(rs!ExcelSheet, "")
        Dim tblName   As String: tblName = Nz(rs!ExcelTable, "")
        
        ' --- [ID:XX] 情報取得 ---
        Debug.Print "--- [ID:" & currentID & "] " & procName & " ---"
        Debug.Print "    Query: " & qryName
        Debug.Print "    Sheet: " & snName
        Debug.Print "    Table: " & tblName
        
        ' (A) SQL更新プロセス
        If qryName <> "" And sqlTemplate <> "" Then
            Dim finalSQL As String
            finalSQL = Replace(sqlTemplate, "{TERM}", targetTerm)
            
            On Error Resume Next
            db.QueryDefs(qryName).sql = finalSQL
            If Err.Number <> 0 Then
                If Err.Number = 3265 Then
                    Debug.Print "    [INFO] クエリが存在しないため新規作成します: " & qryName
                    db.CreateQueryDef qryName, finalSQL
                Else
                    Debug.Print "    [ERROR] クエリ更新失敗 (" & Err.Number & "): " & Err.Description
                    Debug.Print "    [DEBUG] 失敗したSQL: " & Left(finalSQL, 200) & "..."
                End If
                Err.Clear
            Else
                Debug.Print "    [OK] クエリSQLを更新しました。"
            End If
            On Error GoTo Err_Handler
        End If
        
        ' (B) Excelオープンプロセス
        If xlPath <> "" Then
            If Not openWBs.Exists(xlPath) Then
                If Dir(xlPath) <> "" Then
                    Set xlBook = xlApp.Workbooks.Open(xlPath)
                    openWBs.Add xlPath, xlBook
                Else
                    Debug.Print "    [ERROR] Excelファイルが見つかりません: " & xlPath
                    GoTo Next_Record
                End If
            Else
                Set xlBook = openWBs(xlPath)
            End If
        End If
        
        ' (C) シート・テーブル転送プロセス
        Set xlSheet = Nothing
        On Error Resume Next
        Set xlSheet = G_GetSheetByCodeName(xlBook, snName)
        On Error GoTo Err_Handler
        
        If Not xlSheet Is Nothing Then
            Debug.Print "    [EXEC] データを転送します..."
            Call TransferQueryToExcelTable(db, xlSheet, tblName, qryName)
        Else
            Debug.Print "    [ERROR] シートが見つかりません: " & snName
        End If

Next_Record:
        rs.MoveNext
    Loop
    
    ' 5. 後処理
    Dim key As Variant
    For Each key In openWBs.Keys
        Set xlBook = openWBs(key)
        If CLOSE_EXCEL_AFTER_EXPORT Then
            xlBook.Close SaveChanges:=True
        Else
            xlBook.Save
        End If
    Next key
    
    Debug.Print "=== 診断終了 ==="
    MsgBox "プロセスが完了しました。詳細はイミディエイトウィンドウを確認してください。", vbInformation
    
    GoTo Clean_Up

Err_Handler:
    Debug.Print "    [FATAL] 予期せぬエラー発生: " & Err.Description
    MsgBox "エラー: " & Err.Description, vbCritical
    
Clean_Up:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set xlBook = Nothing: Set xlApp = Nothing: Set openWBs = Nothing
End Sub

'----------------------------------------------------------------
' 内部関数 : TransferQueryToExcelTable (インテリジェント・マッピング版)
'----------------------------------------------------------------
Private Sub TransferQueryToExcelTable(ByRef db As DAO.Database, ByRef ws As Object, ByVal tblName As String, ByVal qryName As String)
    Dim rs      As DAO.Recordset
    Dim lo      As Object
    Dim rowCount As Long
    Dim col      As Integer
    Dim rsField  As DAO.Field
    Dim dataArr  As Variant
    Dim mapDict  As Object: Set mapDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set rs = db.OpenRecordset(qryName, dbOpenSnapshot)
    If Err.Number <> 0 Then
        Debug.Print "      [!!] Recordsetオープン失敗: " & Err.Description
        Debug.Print "      [!!] 対象クエリ名: " & qryName
        Err.Clear: Exit Sub
    End If
    
    Set lo = ws.ListObjects(tblName)
    If lo Is Nothing Then
        Debug.Print "      [!!] Excelテーブルが見つかりません: " & tblName
        rs.Close: Exit Sub
    End If
    On Error GoTo 0
    
    ' --- A. データの事前取得と既存データのクリア ---
    ' 0件の場合でも直前のデータが残らないように、まずExcelテーブルをクリアする
    Call ClearListObject_LeaveOneRow(ws, tblName)
    
    If rs.EOF Then
        Debug.Print "      [INFO] 対象レコードが0件です。"
        rs.Close: Exit Sub
    End If
    
    rs.MoveLast: rowCount = rs.recordCount: rs.MoveFirst
    Debug.Print "      [INFO] 対象件数: " & rowCount & " 件"
    
    ' 配列に一括取得 (Fields, Rows) の形式で返る
    dataArr = rs.GetRows(rowCount)
    
    ' --- B. 列マッピングの作成 ---
    ' Excel列名 -> Recordsetのインデックス
    Dim i As Integer
    Debug.Print "      [MAP] タイトル一致を確認中..."
    
    ' デバッグ用：Access側のフィールド名を全て書き出す
    Dim debugFields As String: debugFields = ""
    For i = 0 To rs.Fields.count - 1
        debugFields = debugFields & "[" & rs.Fields(i).Name & "] "
    Next i
    Debug.Print "      [DEBUG] クエリ側の列候補: " & debugFields
    
    For col = 1 To lo.ListColumns.count
        Dim rawTitle As String: rawTitle = lo.ListColumns(col).Name
        Dim exTitle As String: exTitle = Normalize_Text(rawTitle)
        
        For i = 0 To rs.Fields.count - 1
            Dim rsName As String: rsName = rs.Fields(i).Name
            Dim rsNorm As String: rsNorm = Normalize_Text(rsName)
            
            ' 1. 完全一致 (正規化後)
            If rsNorm = exTitle Then
                mapDict.Add col, i
                Exit For
            End If
            
            ' 2. 部分一致 (例: K.期 や at_kenmu.期 に対処)
            ' ドットの後ろ側が一致するか確認
            If InStr(rsNorm, "." & exTitle) > 0 Then
                mapDict.Add col, i
                Exit For
            End If
        Next i
    Next col
    
    rs.Close
    
    If mapDict.count = 0 Then
        Debug.Print "      [WARN] 一致するフィールド名が1つもありませんでした。"
        Exit Sub
    Else
        Debug.Print "      [INFO] マッピング成功: " & mapDict.count & " / " & lo.ListColumns.count & " 列"
    End If
    
    ' --- C. 転送用配列の作成 (Rows, Cols) ---
    Dim exportArr() As Variant
    ReDim exportArr(1 To rowCount, 1 To lo.ListColumns.count)
    
    Dim r As Long
    For r = 1 To rowCount
        Dim key As Variant
        For Each key In mapDict.Keys ' key = Excel列番号
            exportArr(r, key) = dataArr(mapDict(key), r - 1)
        Next key
    Next r
    
    ' --- D. Excelへの一括貼り付け ---
    ' (事前クリアは A の段階で実施済み)

    
    ' テーブルのリサイズ（行数に合わせて拡張）
    Dim startCell As Object: Set startCell = lo.HeaderRowRange.Cells(1, 1)
    lo.Resize ws.Range(startCell, startCell.Offset(rowCount, lo.ListColumns.count - 1))
    
    ' データ貼り付け
    lo.DataBodyRange.Value = exportArr
    Debug.Print "      [SUCCESS] " & mapDict.count & " カラムのデータを同期しました。"
    
End Sub

Private Sub ClearListObject_LeaveOneRow(ByRef ws As Object, ByVal tblName As String)
    Dim lo As Object
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ListRows.count > 1 Then
            ' 2行目以降を一気に削除
            ws.Range(lo.DataBodyRange.rows(2), lo.DataBodyRange.rows(lo.ListRows.count)).Delete
        End If
        ' 1行目のクリア
        lo.DataBodyRange.rows(1).ClearContents
    End If
End Sub

'----------------------------------------------------------------
' 復旧用：管理テーブル (_at_ExportConfig) のSQLとクエリ名を一括正常化する (最強版)
'----------------------------------------------------------------
Public Sub Setup_ExportConfig_Clean()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    
    Set rs = db.OpenRecordset("_at_ExportConfig", dbOpenDynaset)
    
    Do Until rs.EOF
        Select Case rs!ID
            Case 1 ' q_受注_月計_小口
                rs.Edit
                rs!queryName = "q_Exp_Ord_MonSmall"
                rs!sqlTemplate = "SELECT First(基本工事コード) AS [基本工事コード ], First(基本工事名称) AS [基本工事名称 ], Sum(工事価格) AS [工事価格_合計 ], Sum(粗利益額) AS [粗利益額_合計 ], 施工管轄組織名, (受注期 & '期') AS [受注期表示 ], (受注Q & 'Q') AS [受注Q表示 ], (受注月 & '月') AS [受注月表示 ], 基本工事名_官民 FROM at_Icube_累計 WHERE (施工管轄組織名 <> 'ビルサービスグループ') AND (受注期 = Val(Replace('{TERM}','期',''))) AND (所属組織名 = 'ＬＣＳ事業部') AND (一件工事判定 = '小口工事') GROUP BY 施工管轄組織名, 受注期, 受注Q, 受注月, 基本工事名_官民;"
                rs.Update
                
            Case 2 ' q_完工_月計_小口
                rs.Edit
                rs!queryName = "q_Exp_Act_MonSmall"
                rs!sqlTemplate = "SELECT First(基本工事コード) AS [基本工事コード ], First(基本工事名称) AS [基本工事名称 ], Sum(工事価格) AS [工事価格_合計 ], Sum(粗利益額) AS [粗利益額_合計 ], 施工管轄組織名, (完工期 & '期') AS [完工期表示 ], (完工Q & 'Q') AS [完工Q表示 ], (完工月 & '月') AS [完工月表示 ], 基本工事名_官民 FROM at_Icube_累計 WHERE (施工管轄組織名 <> 'ビルサービスグループ') AND (完工期 = Val(Replace('{TERM}','期',''))) AND (所属組織名 = 'ＬＣＳ事業部') AND (一件工事判定 = '小口工事') GROUP BY 施工管轄組織名, 完工期, 完工Q, 完工月, 基本工事名_官民;"
                rs.Update
                
            Case 3 ' q_受注_月計_一件
                rs.Edit
                rs!queryName = "q_Exp_Ord_MonPro"
                rs!sqlTemplate = "SELECT 工事コード, First(工事名称) AS [工事名称 ], Sum(工事価格) AS [工事価格_合計 ], Sum(粗利益額) AS [粗利益額_合計 ], 施工管轄組織名, (受注期 & '期') AS [受注期表示 ], (受注Q & 'Q') AS [受注Q表示 ], (受注月 & '月') AS [受注月表示 ] FROM at_Icube_累計 WHERE (施工管轄組織名 <> 'ビルサービスグループ') AND (受注期 = Val(Replace('{TERM}','期',''))) AND (所属組織名 = 'ＬＣＳ事業部') AND (一件工事判定 = '一件工事') GROUP BY 工事コード, 施工管轄組織名, 受注期, 受注Q, 受注月;"
                rs.Update

            Case 4 ' q_受注_月計_建築
                rs.Edit
                rs!queryName = "q_Exp_Ord_MonConst"
                rs!sqlTemplate = "SELECT First(施工管轄組織名) AS [施工管轄組織名 ], (受注期 & '期') AS [受注期表示 ], (受注Q & 'Q') AS [受注Q表示 ], Sum(工事価格) AS [工事価格_合計 ], Sum(粗利益額) AS [粗利益額_合計 ], First(一件工事判定) AS [一件工事判定 ], First(リニューアル環境区分名) AS [リニューアル環境区分名 ] FROM at_Icube_累計 WHERE (受注期 = Val(Replace('{TERM}','期',''))) AND (リニューアル環境区分名 = 'リニューアル') AND (所属組織名 = '建築部') GROUP BY (受注Q & 'Q');"
                rs.Update

            Case 5, 15, 21 ' 類似SQL
                rs.Edit
                rs!sqlTemplate = "SELECT 施工管轄組織名, (受注期 & '期') AS [受注期表示 ], (受注Q & 'Q') AS [受注Q表示 ], 基本工事名_官民, Sum(工事価格) AS [工事価格_合計 ], Sum(粗利益額) AS [粗利益額_合計 ] FROM at_Icube_累計 WHERE (施工管轄組織名 <> 'ビルサービスグループ') AND (受注期 = Val(Replace('{TERM}','期',''))) AND (所属組織名 = 'ＬＣＳ事業部') AND (一件工事判定 = '小口工事') GROUP BY 施工管轄組織名, 受注期, 受注Q, 基本工事名_官民;"
                rs.Update

            Case 7 ' q_受注_月計_形態別_小口
                rs.Edit
                rs!sqlTemplate = "SELECT 施工管轄組織名, (受注期 & '期') AS [受注期表示 ], (受注Q & 'Q') AS [受注Q表示 ], (受注月 & '月') AS [受注月表示 ], 受注形態名, Sum(工事価格) AS [工事価格_合計 ], Sum(粗利益額) AS [粗利益額_合計 ] FROM at_Icube_累計 WHERE (施工管轄組織名 <> 'ビルサービスグループ') AND (受注期 = Val(Replace('{TERM}','期',''))) AND (所属組織名 = 'ＬＣＳ事業部') AND (一件工事判定 = '小口工事') GROUP BY 施工管轄組織名, 受注期, 受注Q, 受注月, 受注形態名;"
                rs.Update
                
            Case 8 ' q_受注_月計_区分なし
                rs.Edit
                rs!sqlTemplate = "SELECT 施工管轄組織名, (受注期 & '期') AS [受注期表示 ], (受注Q & 'Q') AS [受注Q表示 ], (受注月 & '月') AS [受注月表示 ], Sum(工事価格) AS [工事価格_合計 ], Sum(粗利益額) AS [粗利益額_合計 ] FROM at_Icube_累計 WHERE (施工管轄組織名 <> 'ビルサービスグループ') AND (受注期 = Val(Replace('{TERM}','期',''))) AND (所属組織名 = 'ＬＣＳ事業部') AND (一件工事判定 = '小口工事') GROUP BY 施工管轄組織名, 受注期, 受注Q, 受注月;"
                rs.Update

            Case 11 ' q_受注完工予測_加重平均集計
                rs.Edit
                rs!queryName = "q_Exp_JyuYosoku_Kajyu"
                rs!sqlTemplate = "SELECT * FROM [at_Work_受注完工予測_加重平均集計];"
                rs.Update
                
            Case 12 ' at_Work_受注完工予測_加重平均集計
                rs.Edit
                rs!queryName = "q_Exp_JyuYosoku_Avg"
                rs!sqlTemplate = "SELECT * FROM [at_Work_受注完工予測_加重平均集計] WHERE (予測ターゲット LIKE '{TERM}*');"
                rs.Update

            Case 13 ' at_Work_予測完工高_最終結果
                rs.Edit
                rs!queryName = "q_Exp_ActFcst_Final"
                rs!sqlTemplate = "SELECT * FROM [at_Work_予測完工高_最終結果] WHERE (期_予測ターゲット LIKE '{TERM}*');"
                rs.Update
                
            Case 18 ' q_受注_月計_形態別_小口 (閉じカッコ修正)
                rs.Edit
                rs!sqlTemplate = "SELECT at_Icube_累計.施工管轄組織名, ([受注期] & '期') AS [受注期表示 ], ([受注Q] & 'Q') AS [受注Q表示 ], [受注月] & '月' AS [受注月表示 ], at_Icube_累計.受注形態名, Sum(at_Icube_累計.工事価格) AS [工事価格_合計 ], Sum(at_Icube_累計.粗利益額) AS [粗利益額_合計 ] FROM at_Icube_累計 WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND ([受注期]=Val(Replace('{TERM}','期',''))) AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事')) GROUP BY at_Icube_累計.施工管轄組織名, ([受注期] & '期'), ([受注Q] & 'Q'), [受注月] & '月', at_Icube_累計.受注形態名;"
                rs.Update
                
            Case 21 ' sel_原価S_基本工事
                rs.Edit
                rs!sqlTemplate = "SELECT kt.仮基本工事コード, kt.仮基本工事略名, g.基本工事コード, g.基本工事名, g.工事価格, g.[工事原価(経費込)], g.予定利益, g.粗利率, g.直接工事費, g.経費, g.作業所経費, g.率, g.共通経費, g.率2, g.[既払高：総額], g.[既払高：経費], g.今後支払予定, g.当月より前の支払金額, g.当月支払金額, g.[設計料・他], g.当月以降予定金額, g.行分類 FROM at_原価S_基本工事 AS g INNER JOIN at_Icube_累計 AS kt ON g.基本工事コード = kt.s基本工事コード GROUP BY kt.仮基本工事コード, kt.仮基本工事略名, g.基本工事コード, g.基本工事名, g.工事価格, g.[工事原価(経費込)], g.予定利益, g.粗利率, g.直接工事費, g.経費, g.作業所経費, g.率, g.共通経費, g.率2, g.[既払高：総額], g.[既払高：経費], g.今後支払予定, g.当月より前の支払金額, g.当月支払金額, g.[設計料・他], g.当月以降予定金額, g.行分類;"
                rs.Update
        End Select
        rs.MoveNext
    Loop
    rs.Close
    
    MsgBox "エクスポート設定のクリーンアップ（最強版）が完了しました。" & vbCrLf & "再度エクスポートを実行し、結果を確認してください。", vbInformation
End Sub

'----------------------------------------------------------------
' 診断用：テーブルのフィールド名を一覧表示する
'----------------------------------------------------------------
Public Sub Print_Table_Fields(ByVal tblName As String)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim result As String
    
    On Error Resume Next
    Set tdf = db.TableDefs(tblName)
    If tdf Is Nothing Then
        Debug.Print "[ERROR] テーブルが見つかりません: " & tblName
        Exit Sub
    End If
    
    Debug.Print "--- [" & tblName & "] フィールド名一覧 ---"
    For Each fld In tdf.Fields
        result = result & "[" & fld.Name & "] "
    Next fld
    Debug.Print result
    On Error GoTo 0
End Sub
