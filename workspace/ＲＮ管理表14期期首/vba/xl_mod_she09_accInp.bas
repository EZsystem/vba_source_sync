Attribute VB_Name = "she09_accInp"
'-------------------------------------
' Module: she09_accInp
' 説明  : Accessのクエリ結果をExcelに出力（ヘッダー付き配列を使用）
' 作成日: 2025/05/14
' 修正者: そうじろう
' 更新日: 2025/12/01
'-------------------------------------
Option Explicit

'============================================
' プロシージャ名         : Sub Inport_koujiitiran_FromAccess
' 概要                   : Accessクエリ結果をExcelテーブルに出力する
' 引数                   : なし
' 戻り値（Functionのみ） : -
' 呼び出し元フォーム／イベント : 未特定
' 関連情報               : null値とテーブル空状態のエラーハンドリングを強化
' 備考                   : テーブルが空の場合やnull値に対応
'============================================
Public Sub Inport_koujiitiran_FromAccess()

    ' --- 1. 初期化 ---
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets("tbl")
    Set tbl = ws.ListObjects("tbl_工事一覧")

    ' --- Access接続パラメータの取得 ---
    Dim accPath As String: accPath = ws.Range("B3").value
    Dim queryName As String: queryName = ws.Range("B4").value

    ' --- 2. 条件判定 ---
    If Dir(accPath) = "" Then
        MsgBox "Accessファイルが見つからないにゃ：" & accPath, vbCritical
        Exit Sub
    End If

    ' --- Accessデータ取得（共通クラス使用） ---
    Dim fetcher As New com_clsAccessFetcher
    fetcher.FilePath = accPath

    Dim arrData As Variant
    arrData = fetcher.FetchArray(queryName)
    If IsEmpty(arrData) Then
        MsgBox "データが取得できなかったにゃ", vbExclamation
        Exit Sub
    End If

    ' --- フィールド名（タイトル）取得 ---
    Dim fieldNames() As String
    fieldNames = fetcher.FetchFieldNames(queryName)

    ' --- 3. 実行処理 ---
    ' --- ヘッダー付き配列を構築 ---
    Dim rowCount As Long: rowCount = UBound(arrData, 1)
    Dim colCount As Long: colCount = UBound(arrData, 2)

    Dim arrWithHeader() As Variant
    ReDim arrWithHeader(1 To rowCount + 1, 1 To colCount)

    Dim i As Long, j As Long
    For j = 1 To colCount
        arrWithHeader(1, j) = fieldNames(j)
    Next j

    For i = 1 To rowCount
        For j = 1 To colCount
            ' ※注意：null値をEmpty文字列に変換してエラーを防ぐ
            If IsNull(arrData(i, j)) Then
                arrWithHeader(i + 1, j) = ""
            Else
                arrWithHeader(i + 1, j) = arrData(i, j)
            End If
        Next j
    Next i

    ' --- タイトル取得（テーブルヘッダー全列）with null対応 ---
    Dim titleRng As Range
    Dim arrTitle As Variant
    
    ' テーブルが空の場合の対応
    If tbl.ListRows.Count = 0 Then
        ' ！警告：テーブルが空の場合はHeaderRowRangeのみ使用
        Set titleRng = tbl.HeaderRowRange
    Else
        ' 通常の場合
        Set titleRng = tbl.HeaderRowRange
    End If
    
    ' ヘッダー範囲の値を安全に取得
    If titleRng.Cells.Count = 1 Then
        ' 単一セルの場合
        ReDim arrTitle(1 To 1, 1 To 1)
        If IsNull(titleRng.value) Or IsEmpty(titleRng.value) Then
            arrTitle(1, 1) = ""
        Else
            arrTitle(1, 1) = titleRng.value
        End If
    Else
        ' 複数セルの場合
        arrTitle = titleRng.value
        ' null値チェックと変換
        For j = 1 To UBound(arrTitle, 2)
            If IsNull(arrTitle(1, j)) Or IsEmpty(arrTitle(1, j)) Then
                arrTitle(1, j) = ""
            End If
        Next j
    End If

    ' --- 配列ヘルパーでマッチング処理 ---
    Dim helper As New com_clsArrayHelper
    helper.data = arrWithHeader

    ' --- 出力用配列構築 ---
    Dim outRows As Long: outRows = rowCount
    Dim outCols As Long: outCols = UBound(arrTitle, 2)
    Dim outArr() As Variant
    ReDim outArr(1 To outRows, 1 To outCols)

    For j = 1 To outCols
        Dim idx As Long
        Dim headerValue As String
        
        ' ヘッダー値の安全な取得
        If IsNull(arrTitle(1, j)) Or IsEmpty(arrTitle(1, j)) Then
            headerValue = ""
        Else
            headerValue = CStr(arrTitle(1, j))
        End If
        
        idx = helper.GetColIndex(headerValue)  ' 見出し名で列位置を検索
        If idx > 0 Then
            For i = 1 To outRows
                ' 出力時もnull値チェック
                If IsNull(arrWithHeader(i + 1, idx)) Then
                    outArr(i, j) = ""
                Else
                    outArr(i, j) = arrWithHeader(i + 1, idx)
                End If
            Next i
        Else
            ' マッチしない列は空文字で埋める
            For i = 1 To outRows
                outArr(i, j) = ""
            Next i
        End If
    Next j

    ' --- 4. 結果の出力 ---
    ' --- 一括出力（テーブルのデータ開始位置に合わせる） ---
    Dim accessor As New xl_clsRangeAccessor
    Dim dstTopLeft As Range
    
    ' テーブルが空の場合の対応
    If tbl.ListRows.Count = 0 Then
        ' 新しい行を追加してから出力位置を決定
        tbl.ListRows.Add
        Set dstTopLeft = tbl.DataBodyRange.Cells(1, 1)
        ' 追加した空行を削除（リサイズで対応）
        tbl.Resize tbl.Range.Resize(outRows + 1, tbl.Range.Columns.Count)
    Else
        Set dstTopLeft = tbl.DataBodyRange.Cells(1, 1)
    End If

    accessor.ArrayToRange ws, dstTopLeft.Resize(outRows, outCols), outArr

    'MsgBox "Accessデータの出力が完了したにゃー！", vbInformation

End Sub   ' ← Subの終わり
