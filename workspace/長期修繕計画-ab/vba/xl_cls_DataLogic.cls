VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_DataLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: cls_DataLogic
' 説明　：データの整合性チェック・IDマッピング（種類列無視版）
'-------------------------------------
Option Explicit

'============================================
' メソッド名 : MapBreakdownIDs
' 概要       : 分類テーブルを照合しIDを埋める
' 修正点     : 「種類」列（修繕/更新）を照合キーから除外しました。
'              周期と更新周期の組み合わせだけで特定します。
'============================================
Public Sub MapBreakdownIDs(wsBreakdown As Worksheet, wsCategory As Worksheet)
    
    Dim tblData As ListObject, tblCat As ListObject
    Dim dataArr As Variant, catArr As Variant
    Dim dictMap As Object
    Dim i As Long
    Dim key As String
    Dim res() As Variant
    
    ' --- 1. テーブル取得確認 ---
    If Not GetTable(wsBreakdown, "tbl_内訳", tblData) Then Exit Sub
    If Not GetTable(wsCategory, "tbl_内訳ID", tblCat) Then Exit Sub
    
    ' --- 2. 必須列の存在確認 ---
    ' ※種類列も列としては必要ですが、キーには使いません
    Dim reqCols As Variant
    reqCols = Array("内訳ID", "大分類", "中分類", "周期", "更新周期")
    
    If Not CheckColumnsExist(tblCat, reqCols, "分類シート") Then Exit Sub
    If Not CheckColumnsExist(tblData, reqCols, "内訳シート") Then Exit Sub
    
    ' --- 3. 分類テーブル（マスタ）を辞書にロード ---
    Set dictMap = CreateObject("Scripting.Dictionary")
    dictMap.CompareMode = vbTextCompare
    
    catArr = tblCat.DataBodyRange.value
    
    ' 列インデックス（分類側）
    Dim cID As Long, cL As Long, cM As Long, cC As Long, cU As Long
    With tblCat.ListColumns
        cID = .Item("内訳ID").Index
        cL = .Item("大分類").Index
        cM = .Item("中分類").Index
        cC = .Item("周期").Index
        cU = .Item("更新周期").Index
    End With
    
    For i = 1 To UBound(catArr, 1)
        ' キー生成：大分類 | 中分類 | 周期 | 更新周期
        ' ★種類（cT）は含めない
        key = Join(Array(Trim(catArr(i, cL)), Trim(catArr(i, cM)), Trim(catArr(i, cC)), Trim(catArr(i, cU))), "|")
        dictMap(key) = catArr(i, cID)
    Next i
    
    ' --- 4. 内訳テーブル（ターゲット）を走査 ---
    dataArr = tblData.DataBodyRange.value
    
    ' 列インデックス（内訳側）
    Dim dL As Long, dM As Long, dC As Long, dU As Long
    With tblData.ListColumns
        dL = .Item("大分類").Index
        dM = .Item("中分類").Index
        dC = .Item("周期").Index
        dU = .Item("更新周期").Index
    End With
    
    ReDim res(1 To UBound(dataArr, 1), 1 To 1)
    
    For i = 1 To UBound(dataArr, 1)
        
        ' キー生成：大分類 | 中分類 | 周期 | 更新周期
        ' ★内訳シートに「修繕」と書いてあろうが「更新」と書いてあろうが無視して、
        '   周期の数字が一致すればIDを引くように変更
        key = Join(Array(Trim(dataArr(i, dL)), Trim(dataArr(i, dM)), Trim(dataArr(i, dC)), Trim(dataArr(i, dU))), "|")
        
        If dictMap.Exists(key) Then
            res(i, 1) = dictMap(key)
        Else
            res(i, 1) = "該当無し"
        End If
    Next i
    
    ' 結果書き込み
    tblData.ListColumns("内訳ID").DataBodyRange.value = res
    
End Sub

'============================================
' メソッド名 : CopyIdpaValues
' 概要       : Idpa="b"の行に対し、1行上の「内訳ID」をコピーする
'============================================
Public Sub CopyIdpaValues(wsBreakdown As Worksheet)
    
    Dim tbl As ListObject
    Dim dataArr As Variant
    Dim res() As Variant
    Dim i As Long
    Dim cIdpa As Long, cID As Long
    
    If Not GetTable(wsBreakdown, "tbl_内訳", tbl) Then Exit Sub
    If Not CheckColumnsExist(tbl, Array("Idpa", "内訳ID"), "内訳シート") Then Exit Sub
    
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    dataArr = tbl.DataBodyRange.value
    
    With tbl.ListColumns
        cIdpa = .Item("Idpa").Index
        cID = .Item("内訳ID").Index
    End With
    
    ReDim res(1 To UBound(dataArr, 1), 1 To 1)
    
    ' 1行目はそのまま
    res(1, 1) = dataArr(1, cID)
    
    For i = 2 To UBound(dataArr, 1)
        If LCase(Trim(CStr(dataArr(i, cIdpa)))) = "b" Then
            ' 上の行（処理済みの値）を継承
            res(i, 1) = res(i - 1, 1)
        Else
            res(i, 1) = dataArr(i, cID)
        End If
    Next i
    
    tbl.ListColumns("内訳ID").DataBodyRange.value = res
    
End Sub

'============================================
' メソッド名 : CheckDuplicates
' 概要       : 重複チェック
'============================================
Public Sub CheckDuplicates(wsBreakdown As Worksheet)
    Dim tbl As ListObject
    Dim dataArr As Variant
    Dim res() As Variant
    Dim dictCheck As Object
    Dim i As Long
    Dim key As String
    
    If Not GetTable(wsBreakdown, "tbl_内訳", tbl) Then Exit Sub
    If Not CheckColumnsExist(tbl, Array("Idpa", "種類", "O-ID", "BELCAコード", "チェック"), "内訳シート") Then Exit Sub
    
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    dataArr = tbl.DataBodyRange.value
    ReDim res(1 To UBound(dataArr, 1), 1 To 1)
    
    Set dictCheck = CreateObject("Scripting.Dictionary")
    
    Dim cIdpa As Long, cType As Long, cOID As Long, cBELCA As Long
    With tbl.ListColumns
        cIdpa = .Item("Idpa").Index
        cType = .Item("種類").Index
        cOID = .Item("O-ID").Index
        cBELCA = .Item("BELCAコード").Index
    End With
    
    For i = 1 To UBound(dataArr, 1)
        If LCase(Trim(CStr(dataArr(i, cIdpa)))) = "b" And Trim(CStr(dataArr(i, cType))) = "修繕" Then
            key = Trim(CStr(dataArr(i, cOID))) & "|" & Trim(CStr(dataArr(i, cBELCA)))
            If dictCheck.Exists(key) Then
                res(i, 1) = "×"
            Else
                dictCheck.Add key, True
                res(i, 1) = ""
            End If
        Else
            res(i, 1) = ""
        End If
    Next i
    
    tbl.ListColumns("チェック").DataBodyRange.value = res
End Sub

'============================================
' メソッド名 : UpdateBelcaMax
' 概要       : BELCAコード最大値更新
'============================================
Public Sub UpdateBelcaMax(wsBreakdown As Worksheet, wsBelca As Worksheet)
    Dim tblUch As ListObject, tblBel As ListObject
    Dim dictMax As Object
    Dim arrBel As Variant, arrUch As Variant
    Dim res() As Variant
    Dim i As Long
    Dim code As String, root As String, sfx As Long
    
    If Not GetTable(wsBreakdown, "tbl_内訳", tblUch) Then Exit Sub
    If Not GetTable(wsBelca, "tbl_修繕計画ID表", tblBel) Then Exit Sub
    
    If Not CheckColumnsExist(tblUch, Array("BELCAコード", "BELCAMax"), "内訳シート") Then Exit Sub
    If Not CheckColumnsExist(tblBel, Array("BELCAコード"), "BELCAシート") Then Exit Sub
    
    Set dictMax = CreateObject("Scripting.Dictionary")
    
    If Not tblBel.DataBodyRange Is Nothing Then
        arrBel = tblBel.ListColumns("BELCAコード").DataBodyRange.value
        For i = 1 To UBound(arrBel, 1)
            code = Trim(CStr(arrBel(i, 1)))
            If code <> "" And IsHyphenSecondFromRight(code) Then
                root = Left(code, Len(code) - 2)
                sfx = CLng(Right(code, 1))
                If Not dictMax.Exists(root) Then
                    dictMax(root) = sfx
                Else
                    If sfx > dictMax(root) Then dictMax(root) = sfx
                End If
            End If
        Next i
    End If
    
    If tblUch.DataBodyRange Is Nothing Then Exit Sub
    arrUch = tblUch.ListColumns("BELCAコード").DataBodyRange.value
    ReDim res(1 To UBound(arrUch, 1), 1 To 1)
    
    For i = 1 To UBound(arrUch, 1)
        code = Trim(CStr(arrUch(i, 1)))
        If code <> "" And IsHyphenSecondFromRight(code) Then
            root = Left(code, Len(code) - 2)
            If dictMax.Exists(root) Then
                res(i, 1) = root & "-" & dictMax(root)
            Else
                res(i, 1) = "NoDate"
            End If
        Else
            res(i, 1) = "NoDate"
        End If
    Next i
    
    tblUch.ListColumns("BELCAMax").DataBodyRange.value = res
End Sub

' --- 内部ヘルパー関数 ---

Private Function GetTable(ws As Worksheet, tblName As String, ByRef tblOut As ListObject) As Boolean
    On Error Resume Next
    Set tblOut = ws.ListObjects(tblName)
    On Error GoTo 0
    If tblOut Is Nothing Then
        MsgBox ws.Name & "シートにテーブル '" & tblName & "' が見つかりません。", vbCritical
        GetTable = False
    Else
        GetTable = True
    End If
End Function

Private Function CheckColumnsExist(tbl As ListObject, colNames As Variant, sheetName As String) As Boolean
    Dim colName As Variant
    Dim col As ListColumn
    Dim missingCols As String
    
    For Each colName In colNames
        On Error Resume Next
        Set col = tbl.ListColumns(colName)
        On Error GoTo 0
        If col Is Nothing Then
            missingCols = missingCols & "・" & colName & vbCrLf
        End If
        Set col = Nothing
    Next colName
    
    If missingCols <> "" Then
        MsgBox sheetName & "のテーブルに以下の列が見つかりません：" & vbCrLf & missingCols & vbCrLf & _
               "※列名のスペース有無などを確認してください。", vbCritical
        CheckColumnsExist = False
    Else
        CheckColumnsExist = True
    End If
End Function

Private Function IsHyphenSecondFromRight(s As String) As Boolean
    If Len(s) < 2 Then Exit Function
    If Mid(s, Len(s) - 1, 1) = "-" And IsNumeric(Right(s, 1)) Then
        IsHyphenSecondFromRight = True
    End If
End Function

