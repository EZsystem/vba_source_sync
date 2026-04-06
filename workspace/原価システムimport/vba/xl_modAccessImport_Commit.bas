Attribute VB_Name = "xl_modAccessImport_Commit"
'-------------------------------------
' Module: xl_modAccessImport_Commit
' 説明  : AccessのクエリデータをExcelテーブルに取込む処理（DAO＋配列出力）
' 作成日: 2025/09/03
' 更新日: -
'-------------------------------------

Option Explicit

'============================================
' プロシージャ名: Import_LoadFromAccess
' Module        : xl_modAccessImport_Commit
' 概要          : AccessクエリをDAOで開き、配列に格納してからExcelテーブルに一括書き込みする
' 引数          : なし（シートセル値から参照）
' 戻り値        : なし
'============================================
Public Sub Import_LoadFromAccess()

    ' --- 1. 初期化 ---
    Dim ws As Worksheet
    Dim accPath As String
    Dim queryName As String
    Dim tblName As String
    
    Dim lo As ListObject
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Dim fieldCount As Long, recordCount As Long
    Dim dataArray() As Variant
    Dim headerMap() As Long
    Dim i As Long, r As Long
    
    Set ws = ThisWorkbook.Sheets("原価S_err2")
    accPath = CStr(ws.Range("C4").value)
    queryName = CStr(ws.Range("C5").value)
    tblName = CStr(ws.Range("C6").value)
    
    ' --- 2. テーブル存在確認 ---
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "テーブル '" & tblName & "' が存在しないにゃ", vbExclamation
        Exit Sub
    End If
    
    ' --- 3. Access接続 ---
    Set db = DBEngine.OpenDatabase(accPath)
    Set rs = db.OpenRecordset(queryName, dbOpenSnapshot)
    
    If rs.EOF Then
        ResetTableToOneRow lo
        MsgBox "取込対象データが0件にゃ", vbInformation
        rs.Close: db.Close
        Exit Sub
    End If
    
    ' --- 4. レコード件数・フィールド数を把握 ---
    fieldCount = rs.Fields.Count
    rs.MoveLast
    recordCount = rs.recordCount
    rs.MoveFirst
    
    ' --- 5. ヘッダーとフィールドのマッピング ---
    ReDim headerMap(1 To fieldCount)
    For i = 1 To fieldCount
        headerMap(i) = GetColumnIndexByHeader(lo, rs.Fields(i - 1).Name)
    Next i
    
    ' --- 6. 配列にデータ格納 ---
    ReDim dataArray(1 To recordCount, 1 To lo.ListColumns.Count)
    r = 0
    Do Until rs.EOF
        r = r + 1
        For i = 1 To fieldCount
            If headerMap(i) > 0 Then
                dataArray(r, headerMap(i)) = rs.Fields(i - 1).value
            End If
        Next i
        rs.MoveNext
    Loop
    
    ' --- 7. Excelテーブルに一括出力 ---
    ResetTableToOneRow lo
    If recordCount > 1 Then
        lo.Resize lo.Range.Resize(recordCount + 1, lo.ListColumns.Count)
    End If
    lo.DataBodyRange.value = dataArray
    
    ' --- 8. 終了処理 ---
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
    
    MsgBox "取込完了にゃ（" & recordCount & " 行）", vbInformation
    
End Sub    ' ← Import_LoadFromAccess の終わり


'============================================
' プロシージャ名: ResetTableToOneRow
' Module        : xl_modAccessImport_Commit
' 概要          : テーブルのデータをクリアし、1行だけ空行を残す
'============================================
Private Sub ResetTableToOneRow(ByVal lo As ListObject)
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add AlwaysInsert:=True
    Else
        Do While lo.ListRows.Count > 1
            lo.ListRows(lo.ListRows.Count).Delete
        Loop
        lo.DataBodyRange.Rows(1).ClearContents
    End If
End Sub    ' ← ResetTableToOneRow の終わり


'============================================
' プロシージャ名: GetColumnIndexByHeader
' Module        : xl_modAccessImport_Commit
' 概要          : 指定フィールド名に一致する列番号を返す
'============================================
Private Function GetColumnIndexByHeader(ByVal lo As ListObject, ByVal fieldName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If CStr(lo.HeaderRowRange.Cells(1, i).value) = fieldName Then
            GetColumnIndexByHeader = i
            Exit Function
        End If
    Next i
    GetColumnIndexByHeader = 0
End Function    ' ← GetColumnIndexByHeader の終わり

