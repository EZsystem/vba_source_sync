Attribute VB_Name = "xl_modCostImport_Standalone"
'-------------------------------------
' Module: xl_modAccessImport_Fast
' 説明  : Accessクエリ→Excelテーブル（ListObject）へ高速インポート
' 作成日: 2025/09/03
' 更新日: -
' 参照  : Microsoft ActiveX Data Objects 6.x Library
'         Microsoft Scripting Runtime
'-------------------------------------
Option Explicit

'============================================
' プロシージャ名 : Import_LoadFromAccess_原価S_err2_Fast
' 概要           : シート「原価S_err2」の C4/C5/C6 を参照し、
'                  Accessクエリ結果を Excelテーブルに高速一括出力する
' シート名       : 原価S_err2
' 入力セル       : C4=Accessフルパス, C5=クエリ名, C6=出力テーブル名
' 条件           : 取込前に既存データは1行残してクリア
'                  出力は1行目を含めて上書き、余り行は削除
'============================================
Public Sub Import_LoadFromAccess_原価S_err2_Fast()
    On Error GoTo ErrHandler

    ' --- 0. 表示負荷軽減 ---
    Dim pScreen As Boolean, pCalc As XlCalculation, pEvents As Boolean
    pScreen = Application.ScreenUpdating
    pCalc = Application.Calculation
    pEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' --- 1. 初期化 ---
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dbPath As String, queryName As String, tableName As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset

    Dim fieldDict As Scripting.Dictionary
    Dim i As Long, r As Long, f As Long, tgtCol As Long
    Dim data As Variant, arrField() As String
    Dim dataRows As Long, dataCols As Long
    Dim outArr() As Variant

    Set ws = ThisWorkbook.Sheets("原価S_err2")
    dbPath = CStr(ws.Range("C4").value)
    queryName = CStr(ws.Range("C5").value)
    tableName = CStr(ws.Range("C6").value)

    If Len(dbPath) = 0 Or Len(tableName) = 0 Or Len(queryName) = 0 Then
        MsgBox "C4（DBパス）、C5（クエリ名）、C6（テーブル名）を確認するにゃ", vbExclamation
        GoTo CleanUp
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo ErrHandler
    If lo Is Nothing Then
        MsgBox "指定のテーブルが見つからない: " & tableName, vbCritical
        GoTo CleanUp
    End If

    ' --- 2. ADO接続＆取得 ---
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM [" & queryName & "]", conn, adOpenKeyset, adLockReadOnly

    ' --- 3. データ存在チェック ---
    If rs.EOF Then
        If lo.ListRows.Count = 0 Then lo.ListRows.Add
        lo.DataBodyRange.Rows(1).ClearContents
        GoTo CleanUp
    End If

    ' --- 4. フィールド配列作成（1ベース） ---
    ReDim arrField(1 To rs.Fields.Count)
    For i = 0 To rs.Fields.Count - 1
        arrField(i + 1) = rs.Fields(i).Name
    Next i
    dataCols = rs.Fields.Count

    ' --- 5. データ取得（GetRows：列×行の配列） ---
    data = rs.GetRows()
    dataRows = UBound(data, 2) + 1

    ' --- 6. ヘッダー対応辞書（テーブル見出し→列番号） ---
    Set fieldDict = New Scripting.Dictionary
    fieldDict.CompareMode = TextCompare
    For i = 1 To lo.ListColumns.Count
        fieldDict(lo.HeaderRowRange.Cells(1, i).value) = i
    Next i

    ' --- 7. 出力前クリア（1行残し） ---
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
            .ListRows.Add
        End If
    End With

    ' --- 8. 出力配列構築 ---
    ReDim outArr(1 To dataRows, 1 To lo.ListColumns.Count)
    For r = 1 To dataRows
        For f = 1 To dataCols
            If fieldDict.Exists(arrField(f)) Then
                tgtCol = fieldDict(arrField(f))
                outArr(r, tgtCol) = data(f - 1, r - 1)
            End If
        Next f
    Next r

    ' --- 9. 行数調整＆一括出力 ---
    If lo.ListRows.Count < dataRows Then
        lo.Resize lo.Range.Resize(RowSize:=dataRows + 1)
    End If
    lo.DataBodyRange.Resize(RowSize:=dataRows).value = outArr

    ' --- 10. 余り行削除 ---
    Do While lo.ListRows.Count > dataRows
        lo.ListRows(lo.ListRows.Count).Delete
    Loop

CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    Set rs = Nothing: Set conn = Nothing
    Application.ScreenUpdating = pScreen
    Application.Calculation = pCalc
    Application.EnableEvents = pEvents
    Exit Sub

ErrHandler:
    MsgBox "Import_LoadFromAccess_原価S_err2_Fast でエラー：" & Err.Description, vbCritical
    Resume CleanUp
End Sub

