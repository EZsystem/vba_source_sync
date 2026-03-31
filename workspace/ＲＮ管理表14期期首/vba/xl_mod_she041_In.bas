Attribute VB_Name = "she041_In"
'-------------------------------------
' Module: she041_In
' 説明  : ExcelからExcelへのデータ転写処理（タイトル一致列のみ）
' 作成日: 2025/06/11
' 更新日: -
'-------------------------------------
Option Explicit

'============================================================
' Module: she041_In
' プロシージャ名 : Transfer_AppendToTable
' 概要 : ユーザー選択のブックから、兼務率テーブルにデータ転写する
' 引数 : なし
'============================================================

Public Sub Transfer_AppendToTable()
    Dim srcWb As Workbook
    Dim tgtWs As Worksheet
    Dim srcWs As Worksheet
    Dim tgtTbl As ListObject
    Dim FilePath As String
    Dim srcData As Variant, tgtTitles As Variant, srcTitles As Variant
    Dim matchCols As Object
    Dim i As Long, j As Long, r As Long
    Dim arrOut() As Variant
    Dim lastRow As Long, lastCol As Long

    '-------------------------------------
    ' Step 1 : 転写元ブックを選択する
    '-------------------------------------
    FilePath = Application.GetOpenFilename("Excelファイル (*.xlsx), *.xlsx", , "転写元ファイルを選択してください")
    If FilePath = "False" Then Exit Sub
    
    Set srcWb = Workbooks.Open(FilePath)
    Set srcWs = srcWb.Worksheets(1)

    '-------------------------------------
    ' Step 2 : タイトル・データを読み取る
    '   ・行：従来どおり
    '   ・列：A列〜1行目の最終列までを自動取得
    '-------------------------------------
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).row
    lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

    srcTitles = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(1, lastCol)).value
    srcData = srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(lastRow, lastCol)).value

    '-------------------------------------
    ' Step 3 : 転写先情報取得（テーブルのヘッダー行を使う）
    '-------------------------------------
    Set tgtWs = ThisWorkbook.Worksheets("兼務率")
    Set tgtTbl = tgtWs.ListObjects("t_兼務率RN")
    tgtTitles = tgtTbl.HeaderRowRange.value   '★ここを変更

    '-------------------------------------
    ' Step 4 : タイトル一致列のマッピング
    '-------------------------------------
    Set matchCols = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(tgtTitles, 2)
        For j = 1 To UBound(srcTitles, 2)
            If Trim(CStr(tgtTitles(1, i))) = Trim(CStr(srcTitles(1, j))) Then
                matchCols.Add i, j
                Exit For
            End If
        Next j
    Next i

    If matchCols.Count = 0 Then
        MsgBox "一致するタイトル列が見つかりませんでした", vbExclamation
        srcWb.Close False
        Exit Sub
    End If

    '-------------------------------------
    ' Step 5 : 既存データを1行残してクリア
    '-------------------------------------
    With tgtTbl.DataBodyRange
        If .Rows.Count > 1 Then
            .Offset(1).Resize(.Rows.Count - 1).ClearContents
        End If
    End With

    '-------------------------------------
    ' Step 6 : 一致列のみ配列構成して転写
    '-------------------------------------
    ReDim arrOut(1 To UBound(srcData, 1), 1 To tgtTbl.ListColumns.Count)
    
    For r = 1 To UBound(srcData, 1)
        For i = 1 To tgtTbl.ListColumns.Count
            If matchCols.Exists(i) Then
                arrOut(r, i) = srcData(r, matchCols(i))
            Else
                arrOut(r, i) = "" ' 空欄補完
            End If
        Next i
    Next r

    '-------------------------------------
    ' Step 7 : テーブルに転写
    '-------------------------------------
    tgtTbl.DataBodyRange.Cells(1, 1). _
        Resize(UBound(arrOut, 1), UBound(arrOut, 2)).value = arrOut

    '-------------------------------------
    ' Step 8 : 終了処理
    '-------------------------------------
    srcWb.Close False
    MsgBox "転写が完了しましたニャ！", vbInformation, "そうじろうより"
End Sub


