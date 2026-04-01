Attribute VB_Name = "acc_mod_ExcelExporter"
'----------------------------------------------------------------
' Module: acc_mod_ExcelExporter
' 説明   : 管理テーブル (_at_ExportConfig) 駆動型 Excel エクスポート (診断機能付)
' 更新日 : 2026/03/31
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
        
        Debug.Print "--- [ID:" & currentID & "] " & procName & " ---"
        Debug.Print "    Query: " & qryName
        Debug.Print "    ExcelTable: " & tblName
        
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
                End If
                Err.Clear
            Else
                Debug.Print "    [OK] クエリSQLを更新しました。"
            End If
            On Error GoTo Err_Handler
        Else
            Debug.Print "    [SKIP] クエリ名またはテンプレートが空です。"
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
        Else
            Debug.Print "    [ERROR] Excelパスが指定されていません。"
            GoTo Next_Record
        End If
        
        ' (C) シート・テーブル転送プロセス
        Set xlSheet = Nothing
        On Error Resume Next
        ' シート名（Caption）ではなくオブジェクト名（CodeName）で取得
        Set xlSheet = G_GetSheetByCodeName(xlBook, snName)
        On Error GoTo Err_Handler
        
        If Not xlSheet Is Nothing Then
            ' 転送関数側のログを拾うため、ここで呼び出し
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
' 内部関数 : TransferQueryToExcelTable (診断メッセージ付)
'----------------------------------------------------------------
Private Sub TransferQueryToExcelTable(ByRef db As DAO.Database, ByRef ws As Object, ByVal tblName As String, ByVal qryName As String)
    Dim rs As DAO.Recordset
    Dim lo As Object
    Dim rowCount As Long
    
    ' 1. レコードセット取得
    On Error Resume Next
    Set rs = db.OpenRecordset(qryName, dbOpenSnapshot)
    If Err.Number <> 0 Then
        Debug.Print "      [!!] Recordsetオープン失敗: " & Err.Description
        Err.Clear: Exit Sub
    End If
    
    ' 2. ターゲットテーブルの確認
    Set lo = ws.ListObjects(tblName)
    If lo Is Nothing Then
        Debug.Print "      [!!] Excelテーブルが見つかりません: " & tblName
        rs.Close: Exit Sub
    End If
    On Error GoTo 0
    
    ' 3. テーブルの初期化
    Call ClearListObject_LeaveOneRow(ws, tblName)
    
    ' 4. データの書き込み
    If Not rs.EOF Then
        rs.MoveLast: rowCount = rs.recordCount: rs.MoveFirst
        Dim colCount As Long: colCount = rs.Fields.count
        
        Dim rangeAddr As String
        rangeAddr = lo.HeaderRowRange.Cells(1, 1).Address & ":" & _
                    lo.HeaderRowRange.Cells(rowCount + 1, colCount).Address
        lo.Resize ws.Range(rangeAddr)
        
        lo.DataBodyRange.Cells(1, 1).CopyFromRecordset rs
        Debug.Print "      [SUCCESS] " & rowCount & " 件転送しました。"
    Else
        Debug.Print "      [INFO] 対象レコードが0件です。"
    End If
    
    rs.Close
End Sub

Private Sub ClearListObject_LeaveOneRow(ByRef ws As Object, ByVal tblName As String)
    Dim lo As Object
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ListRows.count > 1 Then
            lo.DataBodyRange.Offset(1, 0).Resize(lo.ListRows.count - 1).Delete
        End If
        lo.DataBodyRange.Rows(1).ClearContents
    End If
End Sub

'----------------------------------------------------------------
' 関数名 : G_GetSheetByCodeName
' 説明   : Workbook内をループし、CodeNameが一致するシートを返します。
'----------------------------------------------------------------
Public Function G_GetSheetByCodeName(ByRef wb As Object, ByVal codeName As String) As Object
    Dim sh As Object
    For Each sh In wb.Sheets
        If sh.codeName = codeName Then
            Set G_GetSheetByCodeName = sh
            Exit Function
        End If
    Next sh
    Set G_GetSheetByCodeName = Nothing
End Function




