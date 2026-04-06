Attribute VB_Name = "she11_acc"
'-------------------------------------
' Module: mod_AccessToExcel_Icube
' 説明  : ExcelシートからAccessクエリを参照し、絞り込んでデータを取得
' 作成日: 2025/06/02
' 更新日: -
'-------------------------------------
Option Explicit


'============================================
' モジュール名　 : she11_acc
' プロシージャ名 : Import_IcubeData
' 概要           : Accessクエリから条件付きでデータ取得し、Excelテーブルに出力
'============================================
Public Sub Import_IcubeData()
    Dim conn As Object ' ADODB.Connection
    Dim rs As Object   ' ADODB.Recordset
    Dim wb As Workbook, ws As Worksheet
    Dim dbPath As String, queryName As String
    Dim sql As String, orgList As String
    Dim period As String
    Dim destTbl As ListObject

    ' --- 対象シートとパラメータ取得 ---
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("I22_Icube加工ALL")
    dbPath = ws.Range("D1").value
    queryName = ws.Range("D2").value
    period = ws.Range("D3").value

    ' --- 所属組織コード（可否 = "○"）のリスト作成 ---
    orgList = GetOrganizationList()

    ' --- SQL組み立て（フィルタ：組織コード + 受注期 or 完工期 >= 指定期） ---
    sql = "SELECT * FROM [" & queryName & "] " & vbCrLf & _
          "WHERE [所属組織コード] IN (" & orgList & ") " & vbCrLf & _
          "AND ([受注期] >= " & period & " OR [完工期] >= " & period & ")"

    ' --- ADO接続 ---
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 1

    ' --- 出力先テーブル（Excel） ---
    Set destTbl = ws.ListObjects("t_受注完工一覧表")

    ' --- データクリア（取込前） ---
    If Not destTbl.DataBodyRange Is Nothing Then
        destTbl.DataBodyRange.Delete
    End If
    
    ' --- データ出力 ---
    If rs.EOF = False Then
        ' テーブルに最低1行だけ追加してから貼る（空テーブル対策）
        destTbl.ListRows.Add
        destTbl.DataBodyRange.Cells(1, 1).CopyFromRecordset rs
    End If


    ' --- データ出力（CopyFromRecordsetで一括） ---
    If Not rs.EOF Then
        destTbl.DataBodyRange.Cells(1, 1).CopyFromRecordset rs
    End If

    ' --- 後処理 ---
    rs.Close: Set rs = Nothing
    conn.Close: Set conn = Nothing

    MsgBox "データの取込が完了したにゃ！", vbInformation
End Sub




'=======================================
' 関数名 : GetOrganizationList
' 概要   : "t_所属組織一覧" シートから「可否 = ○」のコードを取得してIN句文字列にするにゃ
'=======================================
Private Function GetOrganizationList() As String
    Dim tbl As ListObject
    Dim i As Long
    Dim codeCol As Long, flagCol As Long
    Dim result As String

    ' テーブル：t_所属組織一覧（シート：データtbl）
    Set tbl = ThisWorkbook.Sheets("データtbl").ListObjects("t_所属組織一覧")

    ' ヘッダー位置特定
    For i = 1 To tbl.HeaderRowRange.Columns.Count
        If tbl.HeaderRowRange.Cells(1, i).value = "所属組織コード" Then codeCol = i
        If tbl.HeaderRowRange.Cells(1, i).value = "可否" Then flagCol = i
    Next i

    If codeCol = 0 Or flagCol = 0 Then
        MsgBox "「所属組織コード」または「可否」列が見つからないにゃ", vbExclamation
        Exit Function
    End If

    ' 可否 = ○ のコードだけ収集
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(tbl.DataBodyRange.Cells(i, flagCol).value) = "○" Then
            If result <> "" Then result = result & ","
            result = result & "'" & Trim(tbl.DataBodyRange.Cells(i, codeCol).value) & "'"
        End If
    Next i

    GetOrganizationList = result
End Function

