Attribute VB_Name = "she0322_In"
'-------------------------------------
' Module: xl_mod_ImportSonneki
' 説明　：Accessクエリ→Excelテーブルへ高速インポート
' 作成日：2025/06/12
' 更新日：-
'-------------------------------------
Option Explicit

'============================================
' プロシージャ名 : Import_SonnekiData_Fast
' モジュール名   : xl_mod_ImportSonneki
' 概要           : Accessクエリ→Excelテーブルへの高速一括出力（範囲Delete高速化対応）
' シート名       : D2損益期中
' テーブル名     : t_損益収支
'============================================
Sub Import_SonnekiData_Fast()
    On Error GoTo ErrHandler

    ' --- 1. 初期化 ---
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dbPath As String
    Dim queryName As String
    Dim fetcher As com_clsAccessFetcher
    Dim arrData As Variant
    Dim arrField As Variant
    Dim fieldDict As Object
    Dim i As Long, j As Long
    Dim outArr() As Variant
    Dim dataRows As Long, dataCols As Long

    ' --- 2. シート・テーブル取得 ---
    Set ws = ThisWorkbook.Sheets("D2損益期中")
    Set lo = ws.ListObjects("t_損益収支")

    dbPath = ws.Range("E2").value
    queryName = ws.Range("E3").value

    ' --- 3. 既存データ削除（1行だけ残す：範囲Deleteで高速化） ---
    With lo
        If .ListRows.Count > 1 Then
            Dim delRowCount As Long
            delRowCount = .ListRows.Count - 1
            If delRowCount > 0 Then
                Dim delRange As Range
                Set delRange = .DataBodyRange.Rows(2).Resize(delRowCount)
                delRange.Delete xlShiftUp
            End If
        End If
    End With

    ' --- 4. Accessクエリからデータ取得 ---
    Set fetcher = New com_clsAccessFetcher
    fetcher.FilePath = dbPath
    arrData = fetcher.FetchArray(queryName)         ' データ本体
    arrField = fetcher.FetchFieldNames(queryName)    ' フィールド名（1次元）

    If IsEmpty(arrData) Then
        MsgBox "クエリにデータがありません", vbInformation
        GoTo CleanUp
    End If

    dataRows = UBound(arrData, 1)
    dataCols = UBound(arrField)

    ' --- 5. 列名対応辞書を作成（Excelテーブルのヘッダー名と一致した列のみ） ---
    Set fieldDict = CreateObject("Scripting.Dictionary")
    For i = 1 To lo.ListColumns.Count
        fieldDict(lo.HeaderRowRange.Cells(1, i).value) = i
    Next

    ' --- 6. 出力配列を組立 ---
    ReDim outArr(1 To dataRows, 1 To lo.ListColumns.Count)
    For i = 1 To dataRows
        For j = 1 To dataCols
            If fieldDict.Exists(arrField(j)) Then
                outArr(i, fieldDict(arrField(j))) = arrData(i, j)
            End If
        Next
    Next

    ' --- 7. テーブル行を必要数まで増やす ---
    If lo.ListRows.Count < dataRows Then
        lo.Resize lo.Range.Resize(RowSize:=dataRows + 1)
    End If

    ' --- 8. DataBodyRangeに一括書き込み ---
    lo.DataBodyRange.value = outArr

    MsgBox "取込完了", vbInformation

CleanUp:
    Set fetcher = Nothing
    Exit Sub

ErrHandler:
    MsgBox "エラー：" & Err.Description, vbCritical
    Resume CleanUp
End Sub

