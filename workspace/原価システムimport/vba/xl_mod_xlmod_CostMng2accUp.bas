Attribute VB_Name = "xlmod_CostMng2accUp"
'-------------------------------------
' Module: xlmod_CostMng2accUp
' 説明  : Excel配列をAccess仮テーブルへ全件入れ替えアップデート
' 作成日: 2025/06/04
' 更新日: -
'-------------------------------------
Option Explicit

'============================================
' プロシージャ名 : Run_ExportToAccess
' モジュール名   : xlmod_CostMng2accUp
' 概要           : Excel「原価S_temp」シートの配列をAccess仮テーブルへ転送する
' 引数           : なし（E2:ファイルパス, E3:テーブル名を使用）
' 呼び出し先     : ExportArrayToAccessTempTable
' 備考           : タイトルは6行目、データは7行目、B列（列2）から
'============================================
Public Sub Run_ExportToAccess()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("原価S_temp")

    Call ExportArrayToAccessTempTable(ws, 6, 7, 2, ws.Range("C3").value, ws.Range("C4").value) ' (ws, titleRow, dataStartRow, startCol, filePath, tableName)
    MsgBox "処理完了しました"
End Sub

'============================================
' プロシージャ名: ExportArrayToAccessTempTable
' 概要    : Excelの配列データをAccess仮テーブルへアップロード
' 引数    : ws              - データ取得元ワークシート
'         : titleRow        - タイトル行番号（例：6）
'         : dataStartRow    - データ開始行番号（例：7）
'         : startCol        - データ開始列番号（例：2=B列）
'         : filePath        - Accessファイルパスセル（例：Range("E2").value）
'         : tableName       - 出力対象テーブル名セル（例：Range("E3").value）
' 備考    : 仮テーブルは全件削除後、配列で全件追加する
'============================================
Public Sub ExportArrayToAccessTempTable(ws As Worksheet, titleRow As Long, dataStartRow As Long, startCol As Long, FilePath As String, tableName As String)

    ' --- 1. タイトルとデータ範囲の動的取得 ---
    Dim lastCol As Long, lastRow As Long
    lastCol = ws.Cells(titleRow, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).row
    
    Dim titleArr As Variant
    titleArr = ws.Range(ws.Cells(titleRow, startCol), ws.Cells(titleRow, lastCol)).value
    
    Dim dataArr As Variant
    dataArr = ws.Range(ws.Cells(dataStartRow, startCol), ws.Cells(lastRow, lastCol)).value

    ' --- 2. AccessにADOで接続 ---
    Dim cn As Object, cmd As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath

    ' --- 3. テーブルを全件削除 ---
    Dim sqlDelete As String
    sqlDelete = "DELETE FROM [" & tableName & "]"
    cn.Execute sqlDelete

    ' --- 4. テーブル構造とフィールド名配列取得 ---
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & tableName & "] WHERE 1=0", cn, 1, 3 ' adOpenKeyset, adLockOptimistic

    Dim i As Long, fName As String
    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    For i = 1 To rs.Fields.Count
        colMap(rs.Fields(i - 1).Name) = i
    Next

    ' --- 5. タイトル一致したフィールドのみ転写 ---
    Dim titleIndex As Object: Set titleIndex = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(titleArr, 2)
        fName = CStr(titleArr(1, i))
        If colMap.Exists(fName) Then
            titleIndex(fName) = i ' Excel配列上の列番号
        End If
    Next

    ' --- 6. データ追加（1レコードずつ） ---
    Dim r As Long, c As Variant
    For r = 1 To UBound(dataArr, 1)
        rs.AddNew
        For Each c In titleIndex.Keys
            rs.Fields(c).value = dataArr(r, titleIndex(c))
        Next
        rs.Update
    Next

    ' --- 7. 終了処理 ---
    rs.Close: Set rs = Nothing
    cn.Close: Set cn = Nothing

End Sub





