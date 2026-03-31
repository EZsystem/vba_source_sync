Attribute VB_Name = "acc_mod_ExcelExporter_Test"
'Attribute VB_Name = "acc_mod_ExcelExporter_Test"
'----------------------------------------------------------------
' Module: acc_mod_ExcelExporter_Test
' 説明   : Excelテーブルの初期化（1行残し削除）テスト用モジュール
' 更新日 : 2026/03/30
'----------------------------------------------------------------
Option Compare Database
Option Explicit

' エクセルファイルのパス
Private Const EXCEL_PATH As String = "D:\My_code\11_workspaces\RN_kanri_system\RNkanri_system\ＲＮ管理表14期期首.xlsm"
' 対象シート名
Private Const TARGET_SHEET As String = "IcubeData"

' --- Accessクエリ名の定義 ---
Private Const QRY_ICUBE_JYU   As String = "sel_Icube完工月毎リスト_小口工事"
Private Const QRY_ICUBE_KAN   As String = "sel_Icube完工月毎リスト_小口工事"
Private Const QRY_ICUBE_IKEN  As String = "sel_Icube受注月毎リスト_一件工事"
Private Const QRY_ICUBE_KENT  As String = "sel_Icube受注月毎リスト_建築部"
Private Const QRY_GENKA_KIHON As String = "sel_原価S_基本工事"

' --- Excelテーブル名の定義 ---
Private Const TBL_ICUBE_JYU   As String = "xl_IcubeJyu"
Private Const TBL_ICUBE_KAN   As String = "xl_IcubeKan"
Private Const TBL_ICUBE_IKEN  As String = "xl_IcubeIken"
Private Const TBL_ICUBE_KENT  As String = "xl_IcubeKent"
Private Const TBL_GENKA_KIHON As String = "xl_genkaKihon"

' --- Excelシート名の定義 ---
Private Const SHT_ICUBE       As String = "IcubeData"
Private Const SHT_GENKA_EXPORT As String = "原価Data"

'----------------------------------------------------------------
' プロシージャ名 : Execute_Excel_Data_Export
' 概要          : Accessクエリの内容をExcelの各テーブルへ出力する
'----------------------------------------------------------------
Public Sub Execute_Excel_Data_Export()
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim db As DAO.Database: Set db = CurrentDb
    Dim maps As Variant
    Dim i As Integer
    
    On Error GoTo Err_Handler
    
    Debug.Print "--- Excelデータエクスポートを開始します ---"
    
    ' 1. マッピング定義 (Excelシート名, Excelテーブル名, Accessクエリ名)
    maps = Array( _
        Array(SHT_ICUBE, TBL_ICUBE_JYU, QRY_ICUBE_JYU), _
        Array(SHT_ICUBE, TBL_ICUBE_KAN, QRY_ICUBE_KAN), _
        Array(SHT_ICUBE, TBL_ICUBE_IKEN, QRY_ICUBE_IKEN), _
        Array(SHT_ICUBE, TBL_ICUBE_KENT, QRY_ICUBE_KENT), _
        Array(SHT_GENKA_EXPORT, TBL_GENKA_KIHON, QRY_GENKA_KIHON) _
    )
    
    ' 2. Excelアプリケーションの準備
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Open(EXCEL_PATH)
    
    ' 3. 各テーブルへの転送実行
    For i = LBound(maps) To UBound(maps)
        Dim snName  As String: snName = maps(i)(0)
        Dim tblName As String: tblName = maps(i)(1)
        Dim qryName As String: qryName = maps(i)(2)
        
        Set xlSheet = Nothing
        On Error Resume Next
        Set xlSheet = xlBook.Sheets(snName)
        On Error GoTo Err_Handler
        
        If Not xlSheet Is Nothing Then
            Debug.Print "  転送中: [" & qryName & "] -> " & snName & "![" & tblName & "]"
            Call TransferQueryToExcelTable(db, xlSheet, tblName, qryName)
        Else
            Debug.Print "  [!] 警告: シート '" & snName & "' が見つかりません。スキップします。"
        End If
    Next i
    
    ' 4. 完了
    xlBook.Save
    MsgBox "データの流し込みが完了しました。", vbInformation
    
    ' xlBook.Close SaveChanges:=True
    ' xlApp.Quit
    
    Exit Sub

Err_Handler:
    MsgBox "プロセス実行中にエラーが発生しました:" & vbCrLf & Err.Description, vbCritical
    If Not xlBook Is Nothing Then xlBook.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
End Sub

'----------------------------------------------------------------
' 内部関数 : TransferQueryToExcelTable
' 概要    : DAOレコセからListObjectへデータを流し込み、リサイズする
'----------------------------------------------------------------
Private Sub TransferQueryToExcelTable(ByRef db As DAO.Database, ByRef ws As Object, ByVal tblName As String, ByVal qryName As String)
    Dim rs As DAO.Recordset
    Dim lo As Object
    Dim rowCount As Long, colCount As Long
    
    ' 1. レコセ取得
    On Error Resume Next
    Set rs = db.OpenRecordset(qryName, dbOpenSnapshot)
    If Err.Number <> 0 Then
        Debug.Print "    [!] 警告: クエリ '" & qryName & " 'を開けません。"
        Err.Clear: Exit Sub
    End If
    
    ' 2. ターゲットテーブルの確認
    Set lo = ws.ListObjects(tblName)
    If lo Is Nothing Then
        Debug.Print "    [!] 警告: テーブル '" & tblName & "' が存在しません。"
        rs.Close: Exit Sub
    End If
    On Error GoTo 0
    
    ' 3. テーブルの初期化 (1行残し)
    Call ClearListObject_LeaveOneRow(ws, tblName)
    
    ' 4. レコード確認
    If Not rs.EOF Then
        ' レコード件数のカウント (念のためMoveLast)
        rs.MoveLast: rowCount = rs.recordCount: rs.MoveFirst
        colCount = rs.Fields.count
        
        ' 5. テーブルのリサイズ (ヘッダー含む範囲指定)
        ' CopyFromRecordset は自動拡張されない場合があるため、先にResizeで枠を確保
        ' lo.HeaderRowRange は1行、データ行は rowCount 行
        Dim rangeAddr As String
        rangeAddr = lo.HeaderRowRange.Cells(1, 1).Address & ":" & _
                    lo.HeaderRowRange.Cells(rowCount + 1, colCount).Address
        lo.Resize ws.Range(rangeAddr)
        
        ' 6. データの貼り付け (DataBodyRangeの第1セルから)
        lo.DataBodyRange.Cells(1, 1).CopyFromRecordset rs
    End If
    
    rs.Close
    Debug.Print "    -> 完了 (" & rowCount & " 件)"
End Sub

'----------------------------------------------------------------
' プロシージャ名 : Test_ClearExcelTables_Only

'----------------------------------------------------------------
' 内部関数 : ClearListObject_LeaveOneRow
' 概要    : 指定した名前のListObjectを1行残してクリアする
'----------------------------------------------------------------
Private Sub ClearListObject_LeaveOneRow(ByRef ws As Object, ByVal tblName As String)
    Dim lo As Object
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    On Error GoTo 0
    
    If lo Is Nothing Then
        Debug.Print "    [!] 警告: テーブル '" & tblName & "' がシートに存在しません。"
        Exit Sub
    End If
    
    ' データ行が存在するかチェック
    If Not lo.DataBodyRange Is Nothing Then
        ' 2行目以降がある場合は削除
        If lo.ListRows.count > 1 Then
            lo.DataBodyRange.Offset(1, 0).Resize(lo.ListRows.count - 1).Delete
        End If
        ' 1行目の内容をクリア（書式は維持）
        lo.DataBodyRange.Rows(1).ClearContents
    End If
    
    Debug.Print "    -> 初期化完了"
End Sub


