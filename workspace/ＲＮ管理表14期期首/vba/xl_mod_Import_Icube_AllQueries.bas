Attribute VB_Name = "Import_Icube_AllQueries"
Option Explicit

'===========================================================
' プロシージャ名: Import_Icube_AllQueries
' 説明: Accessの4つのクエリ結果をExcelの各テーブルへ転送する
' 共通部品: com_clsAccessFetcher, xl_clsRangeAccessor
'===========================================================
Public Sub Import_Icube_AllQueries()
    Dim dbPath As String: dbPath = "D:\My_DataBase\Icube_.accdb"
    Dim db As Object ' DAO.Database
    
    ' 高速化モードON
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Err_Handler

    ' 1. Accessデータベース接続
    Set db = DBEngine.OpenDatabase(dbPath)

    ' 2. 各クエリとテーブルの処理（クエリ名, テーブル名）
    Call Process_QueryToTable(db, "sel_Icube受注月毎リスト_小口工事", "xl_IcubeJyu")
    Call Process_QueryToTable(db, "sel_Icube完工月毎リスト_小口工事", "xl_IcubeKan")
    Call Process_QueryToTable(db, "sel_Icube受注月毎リスト_一件工事", "xl_IcubeIken")
    Call Process_QueryToTable(db, "sel_Icube受注月毎リスト_建築部", "xl_IcubeKent")

    ' 3. 終了処理
    db.Close
    Set db = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "データの取り込みが完了しましたにゃ！", vbInformation
    Exit Sub

Err_Handler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しましたにゃ: " & Err.Description, vbCritical
    If Not db Is Nothing Then db.Close
End Sub

'===========================================================
' サブプロシージャ: Process_QueryToTable
' 説明: 指定された1つのクエリ結果を項目一致でテーブルへ転送
'===========================================================
Private Sub Process_QueryToTable(ByRef db As Object, ByVal qryName As String, ByVal tblName As String)
    Dim rs As Object ' DAO.Recordset
    Dim lo As ListObject
    Dim fieldNames As Collection
    Dim i As Long, j As Long
    Dim dataArr As Variant
    Dim targetCol As Long
    
    ' Excel側のテーブル取得 (sh_IcubeData オブジェクト名を使用)
    On Error Resume Next
    Set lo = sh_IcubeData.ListObjects(tblName)
    On Error GoTo 0
    
    If lo Is Nothing Then
        Debug.Print "【警告】テーブルが見つかりませんにゃ: " & tblName
        Exit Sub
    End If

    ' --- 1. テーブルの既存データ削除（1行残し） ---
    ' DataBodyRangeが存在する場合のみクリア
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.ClearContents
        ' 行削除が必要な場合は以下を有効化（1行目は残ります）
        If lo.ListRows.Count > 1 Then
            lo.DataBodyRange.Offset(1, 0).Resize(lo.ListRows.Count - 1).Delete
        End If
    End If

    ' --- 2. Accessからデータ取得 ---
    Set rs = db.OpenRecordset(qryName, 2) ' dbOpenSnapshot = 2
    
    If rs.EOF Then
        rs.Close
        Exit Sub
    End If

    ' --- 3. 項目一致マッピング転送 ---
    ' 高速化のため、Recordsetを一括で配列に格納
    dataArr = rs.GetRows(rs.RecordCount)
    ' ※GetRowsは (列, 行) の形式で返るため注意

    ' クエリの各列をループし、Excelテーブルのヘッダーと一致する列へ書き込み
    For i = 0 To rs.Fields.Count - 1
        ' テーブル内の列位置を特定（完全に一致するもののみ）
        targetCol = 0
        On Error Resume Next
        targetCol = WorksheetFunction.Match(rs.Fields(i).Name, lo.HeaderRowRange, 0)
        On Error GoTo 0
        
        If targetCol > 0 Then
            ' 項目が一致した場合、その列のデータを一括転写
            ' GetRowsで取得した配列の該当列（インデックスi）を、テーブルの該当列に配置
            For j = 0 To UBound(dataArr, 2)
                lo.HeaderRowRange.Cells(j + 2, targetCol).value = dataArr(i, j)
            Next j
        End If
    Next i

    rs.Close
    Set rs = Nothing
End Sub

