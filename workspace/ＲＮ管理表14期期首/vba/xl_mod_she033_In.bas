Attribute VB_Name = "she033_In"
'-------------------------------------
' Module: she033_In
' 説明　：Accessクエリ→Excelテーブルへ高速インポート（3処理まとめ）
' 作成日：2025/08/18
' 更新日：-
' 参照　：Microsoft ActiveX Data Objects 6.x Library
'　　　　Microsoft Scripting Runtime
'-------------------------------------
Option Explicit

'============================================
' プロシージャ名 : Import_All_実施表D
' 概要           : 実施表Dシート上の3テーブルを順次取込（高速）
' 対象           : �@tbl_実施表経費（C3, AB4）
'                  �Atbl_実施表設変予定（C3, N4）
'                  �Btbl_実施表工事D（C3, C4）
'============================================
Public Sub Import_All_実施表D()
    On Error GoTo ErrHandler

    ' --- 0. 表示負荷軽減 ---
    Dim pScreen As Boolean, pCalc As XlCalculation, pEvents As Boolean
    pScreen = Application.ScreenUpdating
    pCalc = Application.Calculation
    pEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' --- 1. 取込実行（順不同OK、依存があれば順序調整する） ---
    Import_FromAccess_ToTable_Core _
        sheetName:="D3実施表D", _
        tableName:="tbl_実施表経費", _
        dbPathCell:="C3", _
        queryNameCell:="AB4"

    Import_FromAccess_ToTable_Core _
        sheetName:="D3実施表D", _
        tableName:="tbl_実施表設変予定", _
        dbPathCell:="C3", _
        queryNameCell:="N4"

    Import_FromAccess_ToTable_Core _
        sheetName:="D3実施表D", _
        tableName:="tbl_実施表工事D", _
        dbPathCell:="C3", _
        queryNameCell:="C4"

CleanUp:
    Application.ScreenUpdating = pScreen
    Application.Calculation = pCalc
    Application.EnableEvents = pEvents
    
    Exit Sub

ErrHandler:
    MsgBox "Import_All_実施表Dでエラー：" & Err.Description, vbCritical
    Resume CleanUp
End Sub

'============================================
' プロシージャ名 : Import_FromAccess_ToTable_Core
' 概要           : Accessクエリ→指定ListObjectへ高速一括出力（配列）
' 引数           : sheetName       - シート名
'                  tableName       - ListObject名（テーブル）
'                  dbPathCell      - DBパスのセル番地（例 "C3"）
'                  queryNameCell   - クエリ名のセル番地（例 "AB4"）
' 仕様           : 取込前に既存データを1行残してクリア
'                  出力は1行目から上書き、余り行は削除
'============================================
Public Sub Import_FromAccess_ToTable_Core( _
        ByVal sheetName As String, _
        ByVal tableName As String, _
        ByVal dbPathCell As String, _
        ByVal queryNameCell As String)

    On Error GoTo ErrHandler

    ' --- 1. 初期化 ---
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dbPath As String
    Dim queryName As String
    Dim fetcher As com_clsAccessFetcher
    Dim arrData As Variant            ' 2次元：行×列
    Dim arrField As Variant           ' 1次元：フィールド名
    Dim fieldDict As Scripting.Dictionary
    Dim i As Long, j As Long
    Dim outArr() As Variant
    Dim dataRows As Long, dataCols As Long

    Set ws = ThisWorkbook.Sheets(sheetName)
    Set lo = ws.ListObjects(tableName)

    dbPath = ws.Range(dbPathCell).value
    queryName = ws.Range(queryNameCell).value

    ' --- 2. 既存データ削除（1行だけ残す：範囲Deleteで高速化） ---
    With lo
        If .ListRows.Count > 1 Then
            Dim delRowCount As Long
            delRowCount = .ListRows.Count - 1
            If delRowCount > 0 Then
                Dim delRange As Range
                Set delRange = .DataBodyRange.Rows(2).Resize(delRowCount)
                delRange.Delete xlShiftUp
            End If
        ElseIf .ListRows.Count = 0 Then
            .ListRows.Add    ' 最低1行は確保（後続で上書き）
        End If
    End With

    ' --- 3. Accessクエリから配列取得 ---
    Set fetcher = New com_clsAccessFetcher
    fetcher.FilePath = dbPath
    arrData = fetcher.FetchArray(queryName)          ' 2次元配列（行1..n, 列1..m）
    arrField = fetcher.FetchFieldNames(queryName)    ' 1次元配列（1..m）

    ' --- 4. データなし対応（1行目クリアで終了） ---
    If IsEmpty(arrData) Then
        Dim c As Range
        For Each c In lo.ListRows(1).Range
            c.value = vbNullString
        Next c
        GoTo CleanUp
    End If

    dataRows = UBound(arrData, 1)
    dataCols = UBound(arrField)

    ' --- 5. フィールド名?テーブル列 対応辞書作成 ---
    Set fieldDict = New Scripting.Dictionary
    fieldDict.CompareMode = TextCompare   ' 大文字小文字は無視
    For i = 1 To lo.ListColumns.Count
        fieldDict(lo.HeaderRowRange.Cells(1, i).value) = i
    Next i

    ' --- 6. 出力配列構築（テーブル列数に合わせる。未対応列は空） ---
    ReDim outArr(1 To dataRows, 1 To lo.ListColumns.Count)
    For i = 1 To dataRows
        For j = 1 To dataCols
            If fieldDict.Exists(arrField(j)) Then
                outArr(i, fieldDict(arrField(j))) = arrData(i, j)
            End If
        Next j
    Next i

    ' --- 7. テーブル行数調整（不足分はリサイズ。ヘッダー分+1） ---
    If lo.ListRows.Count < dataRows Then
        lo.Resize lo.Range.Resize(RowSize:=dataRows + 1)
    End If

    ' --- 8. 1行目から上書きで一括出力 ---
    lo.DataBodyRange.Resize(RowSize:=dataRows).value = outArr

    ' --- 9. 旧データの余り行を削除（新データ行数 < 既存行数の場合） ---
    Do While lo.ListRows.Count > dataRows
        lo.ListRows(lo.ListRows.Count).Delete
    Loop

CleanUp:
    Set fetcher = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Import_FromAccess_ToTable_Coreでエラー：" & Err.Description & vbCrLf & _
           "sheet=" & sheetName & ", table=" & tableName, vbCritical
    Resume CleanUp
End Sub


