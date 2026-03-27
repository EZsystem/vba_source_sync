Attribute VB_Name = "she011_Acc"
Option Explicit

Sub she011_AccAll_1()

    Call she011_Acc_cellClear1
    Call she011_Acc_dldate_1

End Sub



Sub she011_Acc_dldate_1()

    ' === 変数宣言 ===
    Dim cn As Object ' ADODB.Connection
    Dim rs As Object ' ADODB.Recordset
    Dim ws As Worksheet ' 対象シート
    Dim dbPath As String ' Accessファイルパス
    Dim tableName As String ' Accessテーブル名
    Dim searchFields As Variant ' 検索対象フィールド名
    Dim searchValue As String ' パラメータ値
    Dim sqlQuery As String ' SQL文
    Dim excelFields As Variant ' Excelの6行目の値（フィールド名）
    Dim validFields As String ' SQL用の有効なフィールド名
    Dim whereClause As String ' WHERE句
    Dim lastCol As Long ' 最終列
    Dim rowIndex As Long ' 出力行のインデックス

    ' === Excelの最適化設定 ===
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False

    ' === シートとAccess設定 ===
    Set ws = ThisWorkbook.Sheets("I22_Icube加工ALL") ' 対象シート
    dbPath = ws.Range("D1").value ' Accessファイルパス
    tableName = ws.Range("D2").value ' Accessテーブル名
    searchFields = Split(ws.Range("D3").value, ",") ' 検索対象フィールド名を配列に分割
    searchValue = ws.Range("D4").value ' パラメータ値

    ' === パラメータ値の確認 ===
    If Trim(searchValue) = "" Then
        MsgBox "パラメータ値が空白です。D4セルを確認してください。", vbExclamation
        Exit Sub
    End If

    ' === 6行目の値を取得（Excelフィールド名） ===
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column ' 6行目の最終列を取得
    excelFields = ws.Range(ws.Cells(6, 1), ws.Cells(6, lastCol)).value ' フィールド名を配列として取得

    ' === SQL用のフィールド名を構築 ===
    validFields = ""
    Dim i As Long
    For i = 1 To UBound(excelFields, 2)
        If validFields <> "" Then validFields = validFields & ", "
        validFields = validFields & "[" & excelFields(1, i) & "]" ' フィールド名を追加
    Next i

    ' === WHERE句を構築（完全一致に変更） ===
    whereClause = ""
    For i = LBound(searchFields) To UBound(searchFields)
        If whereClause <> "" Then whereClause = whereClause & " OR "
        whereClause = whereClause & "[" & Trim(searchFields(i)) & "] = '" & Replace(searchValue, "'", "''") & "'"
    Next i

    ' === SQL文を生成 ===
    sqlQuery = "SELECT " & validFields & " FROM [" & tableName & "]"
    If whereClause <> "" Then
        sqlQuery = sqlQuery & " WHERE " & whereClause
    End If

    ' SQL文をデバッグ出力
    Debug.Print sqlQuery

    ' === Access接続 ===
    On Error GoTo ErrHandler ' エラーハンドリング開始
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    ' === データを取得 ===
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, cn, 1, 1 ' adOpenKeyset, adLockReadOnly

    ' === データをExcelに書き込む ===
    If Not rs.EOF Then
        rowIndex = 7 ' データの出力開始行
        Do Until rs.EOF
            For i = 1 To rs.Fields.Count
                ws.Cells(rowIndex, i).value = rs.Fields(i - 1).value
            Next i
            rowIndex = rowIndex + 1
            rs.MoveNext
        Loop
    Else
        MsgBox "データが見つかりませんでした。", vbExclamation
    End If

    ' === 後処理 ===
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

    ' === Excelの設定を元に戻す ===
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True

    MsgBox "データ取得が完了しました。", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
End Sub


Sub she011_Acc_cellClear1()

    ' === 変数宣言 ===
    Dim ws As Worksheet ' 対象シート
    Dim firstRow As Long, lastRow As Long ' 行範囲
    Dim firstCol As Long, lastCol As Long ' 列範囲

    ' === シート設定 ===
    Set ws = ThisWorkbook.Sheets("I22_Icube加工ALL")

    ' === 範囲の特定 ===
    firstRow = 7 ' 開始行
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row ' B列での最終行を取得
    firstCol = 1 ' A列（1列目）
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column ' 6行での最終列を取得

    ' === 条件確認 ===
    If firstRow >= lastRow Or firstCol >= lastCol Then
        MsgBox "処理対象範囲が無効です。クリア処理をスキップします。", vbExclamation
        Exit Sub
    End If

    ' === 範囲のクリア ===
    ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol)).Clear ' 値と書式をクリア

    ' 完了メッセージ
    'MsgBox "指定範囲のクリアが完了しました。", vbInformation

End Sub
