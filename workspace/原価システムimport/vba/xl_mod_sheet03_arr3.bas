Attribute VB_Name = "sheet03_arr3"
'-------------------------------------
' Module: xl_mod_ErrorChecker_Mother
' 説明  : 原価管理データ突合＆エラー抽出マザー処理
' 作成日: 2025/06/02
' 更新日: -
'-------------------------------------
Option Explicit

'============================================
' プロシージャ名 : sh03_arr3_mother01
' 概要          : 原価管理系データの配列処理＆エラー出力のメイン制御
'============================================
Sub sh03_arr3_mother01()
    ' --- 1. シート範囲クリア ---
    Call sh03_clearRange

    ' --- 2. データ配列取得 ---
    Dim G2_arr As Variant
    Call sh03_arr3_1(G2_arr)

    Dim I22_arr As Variant
    Call sh03_arr3_2(I22_arr)

    Dim I22_arr2 As Variant
    Call sh03_arr3_3(I22_arr, I22_arr2)

    ' --- 3. 原価管理データ突合（取込み漏れ・逆取込漏れ） ---
    Dim Checkarr1 As Variant
    Call sh03_arr3_4(I22_arr2, G2_arr, Checkarr1)

    Dim Checkarr2 As Variant
    Call sh03_arr3_6(I22_arr2, G2_arr, Checkarr2)

    ' --- 4. シート出力 ---
    Call sh03_arr3_55         ' タイトル行追加
    Call sh03_arr3_5(Checkarr1)
    Call sh03_arr3_75         ' タイトル行追加
    Call sh03_arr3_7(Checkarr2)
End Sub

'============================================
' プロシージャ名: sh03_clearRange
' 概要: 出力先シートの出力範囲をクリア
'============================================
Sub sh03_clearRange()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("G3_原価Sエラー調査")
    Dim dataRange As Range, lastCell As Range, clearRange As Range

    On Error Resume Next ' 範囲エラー対応
    Set dataRange = ws.Range("A7").CurrentRegion
    If dataRange Is Nothing Then Exit Sub

    Set lastCell = dataRange.Cells(dataRange.Rows.Count, dataRange.Columns.Count)
    If lastCell.row > 7 Then
        Set clearRange = ws.Range("A8", lastCell)
        clearRange.ClearContents
    End If
    On Error GoTo 0
End Sub

'============================================
' プロシージャ名: sh03_arr3_1
' 概要: G2_原価S加工データの全データを配列化
' 引数: G2_arr - 結果格納（ByRef）
'============================================
Sub sh03_arr3_1(ByRef G2_arr As Variant)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("G2_原価S加工データ")
    Dim firstRow As Long: firstRow = 6
    Dim firstCol As Long: firstCol = 1
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastCol = ws.Cells(7, ws.Columns.Count).End(xlToLeft).Column
    G2_arr = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol)).value
End Sub

'============================================
' プロシージャ名: sh03_arr3_2
' 概要: I22_Icube加工ALLの全データを配列化
' 引数: I22_arr - 結果格納（ByRef）
'============================================
Sub sh03_arr3_2(ByRef I22_arr As Variant)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("I22_Icube加工ALL")
    Dim firstRow As Long: firstRow = 6
    Dim firstCol As Long: firstCol = 1
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    I22_arr = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol)).value
End Sub

'============================================
' プロシージャ名: sh03_arr3_3
' 概要: タイトル行名で条件判定し、対象外データのみ新配列へ格納
' 引数: I22_arr  - 元配列（ByRef）
'       I22_arr2 - 結果配列（ByRef）
' 条件: ・一件工事判定="一件工事" または 所属組織名="建築部" を含む行を除外
'============================================
Sub sh03_arr3_3(ByRef I22_arr As Variant, ByRef I22_arr2 As Variant)
    Dim j As Long, i As Long
    Dim colJudge1 As Long, colJudge2 As Long
    Dim colName1 As String: colName1 = "一件工事判定"
    Dim colValue1 As String: colValue1 = "一件工事"
    Dim colName2 As String: colName2 = "所属組織名"
    Dim colValue2 As String: colValue2 = "建築部"

    ' --- タイトル行インデックス取得 ---
    For j = 1 To UBound(I22_arr, 2)
        If Trim(CStr(I22_arr(1, j))) = colName1 Then colJudge1 = j
        If Trim(CStr(I22_arr(1, j))) = colName2 Then colJudge2 = j
    Next j

    If colJudge1 = 0 Or colJudge2 = 0 Then
        MsgBox "必要なタイトル列が見つかりません: [" & colName1 & "] or [" & colName2 & "]", vbExclamation
        ReDim I22_arr2(1 To 1, 1 To 1): I22_arr2(1, 1) = "タイトルエラー"
        Exit Sub
    End If

    ' --- 一時配列（最大数で用意） ---
    Dim tempArr() As Variant
    ReDim tempArr(1 To UBound(I22_arr, 1), 1 To UBound(I22_arr, 2))
    Dim matchCount As Long: matchCount = 1

    ' タイトル行コピー
    For j = 1 To UBound(I22_arr, 2)
        tempArr(1, j) = I22_arr(1, j)
    Next j

    ' --- 条件判定＆抽出 ---
    For i = 2 To UBound(I22_arr, 1)
        If Trim(CStr(I22_arr(i, colJudge1))) <> colValue1 _
        And Trim(CStr(I22_arr(i, colJudge2))) <> colValue2 Then
            matchCount = matchCount + 1
            For j = 1 To UBound(I22_arr, 2)
                tempArr(matchCount, j) = I22_arr(i, j)
            Next j
        End If
    Next i

    ' --- 結果出力 ---
    If matchCount > 1 Then
        ReDim I22_arr2(1 To matchCount, 1 To UBound(I22_arr, 2))
        For i = 1 To matchCount
            For j = 1 To UBound(I22_arr, 2)
                I22_arr2(i, j) = tempArr(i, j)
            Next j
        Next i
    Else
        ReDim I22_arr2(1 To 1, 1 To UBound(I22_arr, 2))
        I22_arr2(1, 1) = "一致するデータなし"
    End If
End Sub

'============================================
' プロシージャ名: sh03_arr3_4
' 概要: G2_arr(1列目)に存在しないI22_arr2(1列目)の行を抽出
' 引数: I22_arr2, G2_arr, Checkarr1
'============================================
Sub sh03_arr3_4(ByRef I22_arr2 As Variant, ByRef G2_arr As Variant, ByRef Checkarr1 As Variant)
    Dim i As Long, j As Long, matchCount As Long
    Dim dictG2 As Object: Set dictG2 = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(G2_arr, 1)
        If Not dictG2.Exists(G2_arr(i, 1)) Then dictG2.Add G2_arr(i, 1), Nothing
    Next i

    matchCount = 0
    For i = 1 To UBound(I22_arr2, 1)
        If Not dictG2.Exists(I22_arr2(i, 1)) Then matchCount = matchCount + 1
    Next i

    If matchCount > 0 Then
        ReDim Checkarr1(1 To matchCount, 1 To UBound(I22_arr2, 2))
        matchCount = 0
        For i = 1 To UBound(I22_arr2, 1)
            If Not dictG2.Exists(I22_arr2(i, 1)) Then
                matchCount = matchCount + 1
                For j = 1 To UBound(I22_arr2, 2)
                    Checkarr1(matchCount, j) = I22_arr2(i, j)
                Next j
            End If
        Next i
    Else
        ReDim Checkarr1(1 To 1, 1 To 1): Checkarr1(1, 1) = "原価管理への取込み忘れ無し"
    End If
End Sub

'============================================
' プロシージャ名: sh03_arr3_5
' 概要: Checkarr1配列をシート出力（最大5列）
'============================================
Sub sh03_arr3_5(ByRef Checkarr1 As Variant)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("G3_原価Sエラー調査")
    Dim i As Long, j As Long
    Dim startRow As Long: startRow = 8
    Dim startCol As Long: startCol = 1

    If UBound(Checkarr1, 1) = 1 And Checkarr1(1, 1) = "原価管理への取込み忘れ無し" Then
        ws.Cells(startRow, startCol).value = Checkarr1(1, 1)
    Else
        For i = 1 To UBound(Checkarr1, 1)
            For j = 1 To 5
                On Error Resume Next
                ws.Cells(startRow + i - 1, startCol + j - 1).value = Checkarr1(i, j)
                On Error GoTo 0
            Next j
        Next i
    End If
End Sub

'============================================
' プロシージャ名: sh03_arr3_6
' 概要: G2_arrの「枝番工事コード」に存在しないI22_arr2の行をCheckarr2に抽出（タイトル名で条件一致判定）
' 引数: I22_arr2, G2_arr, Checkarr2
'============================================
Sub sh03_arr3_6(ByRef I22_arr2 As Variant, ByRef G2_arr As Variant, ByRef Checkarr2 As Variant)
    Dim i As Long, j As Long, matchCount As Long
    Dim dictG2 As Object: Set dictG2 = CreateObject("Scripting.Dictionary")
    Dim colKey_I22 As Long, colKey_G2 As Long
    Dim colName As String: colName = "枝番工事コード"

    ' --- タイトル行（1行目）から列番号を特定 ---
    For j = 1 To UBound(I22_arr2, 2)
        If Trim(CStr(I22_arr2(1, j))) = colName Then colKey_I22 = j
    Next j
    For j = 1 To UBound(G2_arr, 2)
        If Trim(CStr(G2_arr(1, j))) = colName Then colKey_G2 = j
    Next j

    If colKey_I22 = 0 Or colKey_G2 = 0 Then
        MsgBox "「枝番工事コード」の列が見つからなかったにゃ", vbExclamation
        Exit Sub
    End If

    ' --- G2_arrのコード一覧を辞書に格納（タイトル行はスキップ） ---
    For i = 2 To UBound(G2_arr, 1)
        If Not dictG2.Exists(G2_arr(i, colKey_G2)) Then
            dictG2.Add G2_arr(i, colKey_G2), Nothing
        End If
    Next i

    ' --- 該当しない行をカウント（I22_arr2の中から） ---
    matchCount = 0
    For i = 2 To UBound(I22_arr2, 1)
        If Not dictG2.Exists(I22_arr2(i, colKey_I22)) Then
            matchCount = matchCount + 1
        End If
    Next i

    ' --- 抽出してCheckarr2に格納（タイトル行付き） ---
    If matchCount > 0 Then
        ReDim Checkarr2(1 To matchCount + 1, 1 To UBound(I22_arr2, 2))

        ' タイトル行をコピー
        For j = 1 To UBound(I22_arr2, 2)
            Checkarr2(1, j) = I22_arr2(1, j)
        Next j

        matchCount = 1
        For i = 2 To UBound(I22_arr2, 1)
            If Not dictG2.Exists(I22_arr2(i, colKey_I22)) Then
                matchCount = matchCount + 1
                For j = 1 To UBound(I22_arr2, 2)
                    Checkarr2(matchCount, j) = I22_arr2(i, j)
                Next j
            End If
        Next i
    Else
        ReDim Checkarr2(1 To 1, 1 To 1)
        Checkarr2(1, 1) = "原価管理への工事取込みモレ無し"
    End If
End Sub

'============================================
' プロシージャ名: sh03_arr3_7
' 概要: Checkarr2配列をシート出力（最大6列・タイトル行以降に追記）
'============================================
Sub sh03_arr3_7(ByRef Checkarr2 As Variant)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("G3_原価Sエラー調査")
    Dim i As Long, j As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1

    If UBound(Checkarr2, 1) = 1 And Checkarr2(1, 1) = "原価管理への工事取込みモレ無し" Then
        ws.Cells(lastRow, 1).value = Checkarr2(1, 1)
    Else
        For i = 1 To UBound(Checkarr2, 1)
            For j = 1 To 6
                On Error Resume Next
                ws.Cells(lastRow + i - 1, j).value = Checkarr2(i, j)
                On Error GoTo 0
            Next j
        Next i
    End If
End Sub

'============================================
' プロシージャ名: sh03_arr3_55
' 概要: タイトル（原価管理システムデータ取込み）を最終行へ追記
'============================================
Sub sh03_arr3_55()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("G3_原価Sエラー調査")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    ws.Cells(lastRow + 1, 1).value = "当ファイルへの原価管理システムデータ取込み"
End Sub

'============================================
' プロシージャ名: sh03_arr3_75
' 概要: タイトル（データ取込みモレ）を最終行へ追記
'============================================
Sub sh03_arr3_75()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("G3_原価Sエラー調査")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    ws.Cells(lastRow + 1, 1).value = "原価管理へのデータ取込みモレ"
End Sub


