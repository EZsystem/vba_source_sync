Attribute VB_Name = "she021_Acc"
Option Explicit

'==========================================================
' Sub名: she021_Acc_update_1
' 概要: ExcelのデータをAccessの特定のテーブルに追加し、
'       空白セルをAccessのデフォルト値で補完する
' 日付: 2025/01/27
'==========================================================
Sub she021_Acc_update_1()

    ' 変数宣言
    Dim cn As Object ' ADODB.Connection
    Dim rs As Object ' ADODB.Recordset
    Dim defaultsRs As Object ' ADODB.Recordset for default values
    Dim dbPath As String ' Accessファイルパス
    Dim tableName As String ' Accessテーブル名
    Dim defaultsTable As String ' デフォルト値テーブル名
    Dim ws As Worksheet ' 対象シート
    Dim data As Variant ' Excelデータを格納する配列
    Dim defaults As Object ' デフォルト値を格納するDictionary
    Dim lastRow As Long, lastCol As Long ' 最終行・最終列
    Dim i As Long, j As Long ' ループ用
    Dim fieldName As String ' フィールド名
    Dim cellValue As Variant ' 各セルの値

    ' === Excelの最適化設定 ===
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False

    ' === シート設定 ===
    Set ws = ThisWorkbook.Sheets("G22_原価S基本工事")
    dbPath = ThisWorkbook.Sheets("G1_原価S直データ").Range("R1").value ' Accessファイルパス
    tableName = ThisWorkbook.Sheets("G1_原価S直データ").Range("S2").value ' テーブル名
    defaultsTable = "tb_Excelデフォルト値" ' デフォルト値が格納されているAccessのテーブル名

    ' === データベース接続設定 ===
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    ' === デフォルト値を取得 ===
    Set defaultsRs = CreateObject("ADODB.Recordset")
    defaultsRs.Open "SELECT タイトル名, デフォルト値 FROM " & defaultsTable, cn, 1, 1 ' adOpenStatic, adLockReadOnly

    ' デフォルト値をDictionaryに格納
    Set defaults = CreateObject("Scripting.Dictionary")
    Do While Not defaultsRs.EOF
        defaults.Add defaultsRs.Fields("タイトル名").value, defaultsRs.Fields("デフォルト値").value
        defaultsRs.MoveNext
    Loop
    defaultsRs.Close

    ' === Excelデータを配列に格納 ===
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    data = ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol)).value ' データを配列に格納

    ' === レコードセットを開く ===
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open tableName, cn, 1, 3 ' adOpenKeyset, adLockOptimistic

    ' === データをレコードセットに追加 ===
    For i = 2 To UBound(data, 1) ' データは7行目から開始（タイトル行は6行目）

        rs.AddNew ' 新しいレコードを追加
        For j = 1 To UBound(data, 2) ' 列ごと
            fieldName = data(1, j) ' 配列の1行目がフィールド名（タイトル）
            cellValue = data(i, j) ' 配列の値を取得

            ' 空白セルをデフォルト値で補完
            If IsEmpty(cellValue) Or cellValue = "" Then
                If defaults.Exists(fieldName) Then
                    cellValue = defaults(fieldName) ' デフォルト値を取得
                End If
            End If

            ' フィールドが存在する場合のみ値をセット
            If FieldExists(rs, fieldName) Then
                rs.Fields(fieldName) = cellValue
            End If
        Next j
        rs.Update ' レコードを保存

    Next i

    ' === 後処理 ===
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

    ' === Excelの設定を元に戻す ===
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True

    'MsgBox "Accessテーブルの更新が完了しました。", vbInformation

End Sub

' === フィールド存在確認関数 ===
Function FieldExists(rs As Object, fieldName As String) As Boolean
    On Error Resume Next
    FieldExists = Not IsNull(rs.Fields(fieldName).Name)
    On Error GoTo 0
End Function

