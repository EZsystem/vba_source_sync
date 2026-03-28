Attribute VB_Name = "xlmod_CostMng1"
'-------------------------------------
' Module: xlmod_CostMng1
' 説明  : 原価データの加工処理し、仮テーブルへの出力
' 作成日: 2025/06/02
' 更新日: -
'-------------------------------------
Option Explicit


'============================================
' プロシージャ名 : ArrangeGencaDataMain
' Module        : xl_mod_CostMng1
' 概要          : G1_原価S直データを各種加工後、G2_原価S加工データに出力
' 引数          : なし
' 呼び出し元    : 手動実行または他プロシージャ
'============================================
Public Sub ArrangeGencaDataMain()
    Dim arr1 As Variant

    '--- 1. データ元シートから原価データを配列として取得
    Call LoadGencaSource(arr1)

    '--- 2. 基本工事情報行へ小規模工事名と既払高：経費に値を転写
    Call Transfer_CopyValues_UpDown(arr1)

    '--- 3. 小規模工事名から基本工事コードと基本工事名に転写する
    Call Transfer_ProcessedConstructionName_Updated(arr1)

    '--- 4. 小規模工事名から工事コードと工事名へ転写する
    Call Transfer_ConstructionCodeAndName_Final(arr1)

    '--- 5. 枝番工事コード列において、カウント値を記入する
    Call Count_ResetOnBlank_ThroughAll(arr1)

    '--- 6. 工事コードと枝番工事コードを連結して枝番工事コード列に書き戻す
    Call Concat_ProjectCodeAndBranch(arr1)

    '--- 7. 枝番工事コードのある行に小規模工事名を追加工事名称へ転写
    Call Transfer_AdditionalWorkName(arr1)

    '--- 8. 基本工事コードまたは枝番工事コードが空白の行を削除（№=2は除外）
    Call RemoveRowsByEmptyCode_Final(arr1)

    '--- 9. 工事コード有無で「行分類」列を分岐
    Call ClassifyRowsByConstructionCode(arr1)

    '---10. 「状況」列で値が「予定」のレコードを削除
    Call RemoveRowsByStatus_Yotei(arr1)

    '---11. 工事コードがnullの行で指定列を集計
    Call SummarizeByNullConstructionCode(arr1)

    '---12. 配列をテーブルに上書き出力
    Call OutputArrayToTable_Overwrite(arr1)
    MsgBox "処理完了しました"
End Sub

'============================================
' プロシージャ名 : Concat_ProjectCodeAndBranch
' 概要          : 工事コードと枝番を連結する（№列はまだ触らない）
'============================================
Public Sub Concat_ProjectCodeAndBranch(arr1 As Variant)
    Dim rowCount As Long: rowCount = UBound(arr1, 1)
    Dim colBranch As Long, colMain As Long, r As Long
    Dim codeMain As String, codeBranch As String

    colMain = GetTitleColumn(arr1, "工事コード")
    colBranch = GetTitleColumn(arr1, "枝番工事コード")

    If colMain = 0 Or colBranch = 0 Then
        MsgBox "列が見つからないにゃ", vbExclamation
        Exit Sub
    End If

    For r = 2 To rowCount
        codeBranch = Trim(arr1(r, colBranch) & "")
        If codeBranch <> "" Then
            codeMain = Trim(arr1(r, colMain) & "")
            If codeMain <> "" Then
                ' ここでは連結のみ実行
                arr1(r, colBranch) = codeMain & "-" & codeBranch
            End If
        End If
    Next r
End Sub




'============================================
' プロシージャ名 : SummarizeByNullConstructionCode
' Module        : xl_mod_CostMng1
' 概要          : 工事コードがnullまたは空白の行に対し、次のnullまでの範囲で指定列を集計
' 引数          : arr1 (ByRef Variant) - 処理対象の配列
' 呼び出し元    : ArrangeGencaDataMain
'============================================
Private Sub SummarizeByNullConstructionCode(ByRef arr1 As Variant)
    Dim i As Long, j As Long
    Dim constructionCodeCol As Long
    Dim targetCols As Variant
    Dim targetColIndices() As Long
    Dim summaryStartRow As Long
    Dim summaryEndRow As Long
    Dim colIndex As Long
    Dim sumValue As Double
    Dim cellValue As Variant
    
    ' 対象列名を定義
    targetCols = Array("工事価格", "工事原価(経費込)", "予定利益", "直接工事費", "経費", "作業所経費")
    
    ' 工事コード列のインデックスを検索
    constructionCodeCol = -1
    For j = LBound(arr1, 2) To UBound(arr1, 2)
        If arr1(1, j) = "工事コード" Then
            constructionCodeCol = j
            Exit For
        End If
    Next j
    
    ' 工事コード列が見つからない場合は処理を終了
    If constructionCodeCol = -1 Then
        MsgBox "「工事コード」列が見つかりません。処理を中断します。", vbCritical
        End
    End If
    
    ' 対象列のインデックスを取得
    ReDim targetColIndices(0 To UBound(targetCols))
    For i = 0 To UBound(targetCols)
        targetColIndices(i) = -1
        For j = LBound(arr1, 2) To UBound(arr1, 2)
            If arr1(1, j) = targetCols(i) Then
                targetColIndices(i) = j
                Exit For
            End If
        Next j
        
        ' 対象列が見つからない場合は処理を継続（警告のみ）
        If targetColIndices(i) = -1 Then
            Debug.Print "警告: 「" & targetCols(i) & "」列が見つかりません。"
        End If
    Next i
    
    ' 工事コードがnullまたは空白の行を検索し、集計処理を実行
    For i = 2 To UBound(arr1, 1) ' 2行目から開始（1行目はヘッダー）
        
        ' 工事コードがnullまたは空白かチェック
        If IsNullOrEmpty(arr1(i, constructionCodeCol)) Then
            
            ' 集計範囲の開始行を設定（nullの次の行）
            summaryStartRow = i + 1
            
            ' 集計範囲の終了行を検索（次のnullの前の行、または最終行）
            summaryEndRow = UBound(arr1, 1) ' デフォルトは最終行
            For j = summaryStartRow To UBound(arr1, 1)
                If IsNullOrEmpty(arr1(j, constructionCodeCol)) Then
                    summaryEndRow = j - 1
                    Exit For
                End If
            Next j
            
            ' 集計範囲が有効な場合のみ処理
            If summaryStartRow <= summaryEndRow Then
                
                ' 各対象列について集計
                For colIndex = 0 To UBound(targetColIndices)
                    
                    ' 対象列が存在する場合のみ処理
                    If targetColIndices(colIndex) <> -1 Then
                        
                        sumValue = 0
                        
                        ' 集計範囲の各行を処理
                        For j = summaryStartRow To summaryEndRow
                            cellValue = arr1(j, targetColIndices(colIndex))
                            
                            ' 数値チェック
                            If Not IsEmpty(cellValue) And Not IsNull(cellValue) And cellValue <> "" Then
                                If IsNumeric(cellValue) Then
                                    sumValue = sumValue + CDbl(cellValue)
                                Else
                                    ' 数値以外の値が見つかった場合はエラー
                                    MsgBox "エラー: " & (j) & "行目の「" & targetCols(colIndex) & "」列に数値以外の値が含まれています。" & vbCrLf & _
                                           "値: " & CStr(cellValue) & vbCrLf & "処理を中断します。", vbCritical
                                    End
                                End If
                            End If
                        Next j
                        
                        ' 集計結果をnull行に格納
                        arr1(i, targetColIndices(colIndex)) = sumValue
                        
                    End If
                Next colIndex
            End If
        End If
    Next i
End Sub

'============================================
' 関数名   : IsNullOrEmpty
' 概要     : 値がNull、Empty、または空文字列かを判定
' 引数     : value (Variant) - 判定対象の値
' 戻り値   : Boolean - True:Null/Empty/空文字, False:値あり
'============================================
Private Function IsNullOrEmpty(value As Variant) As Boolean
    IsNullOrEmpty = (IsNull(value) Or IsEmpty(value) Or value = "")
End Function


'============================================
' プロシージャ名 : RemoveRowsByStatus_Yotei
' Module        : xl_mod_CostMng1
' 概要          : 「状況」列で値が「予定」のレコードを削除
' 引数          : arr1 (ByRef Variant) - 処理対象の配列
' 呼び出し元    : ArrangeGencaDataMain
'============================================
Private Sub RemoveRowsByStatus_Yotei(ByRef arr1 As Variant)
    Dim i As Long, j As Long
    Dim statusCol As Long
    Dim newArr As Variant
    Dim validRowCount As Long
    Dim currentRow As Long
    
    ' 「状況」列のインデックスを検索
    statusCol = -1
    For j = LBound(arr1, 2) To UBound(arr1, 2)
        If arr1(1, j) = "状況" Then
            statusCol = j
            Exit For
        End If
    Next j
    
    ' 「状況」列が見つからない場合は処理を終了
    If statusCol = -1 Then
        Exit Sub
    End If
    
    ' 有効な行数をカウント（「予定」以外の行）
    validRowCount = 1 ' ヘッダー行は必ず含める
    For i = 2 To UBound(arr1, 1)
        If arr1(i, statusCol) <> "予定" Then
            validRowCount = validRowCount + 1
        End If
    Next i
    
    ' 新しい配列を作成
    ReDim newArr(1 To validRowCount, LBound(arr1, 2) To UBound(arr1, 2))
    
    ' ヘッダー行をコピー
    For j = LBound(arr1, 2) To UBound(arr1, 2)
        newArr(1, j) = arr1(1, j)
    Next j
    
    ' 「予定」以外の行をコピー
    currentRow = 2
    For i = 2 To UBound(arr1, 1)
        If arr1(i, statusCol) <> "予定" Then
            For j = LBound(arr1, 2) To UBound(arr1, 2)
                newArr(currentRow, j) = arr1(i, j)
            Next j
            currentRow = currentRow + 1
        End If
    Next i
    
    ' 元の配列を新しい配列で置き換え
    arr1 = newArr
End Sub


'============================================
' プロシージャ名 : LoadGencaSource
' Module        : xl_mod_CostMng1
' 概要          : 「G1_原価S直データ」からデータを配列取得
' 引数          : arrOut 配列（ByRef）
'============================================
Private Sub LoadGencaSource(ByRef arrOut As Variant)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Set ws = ThisWorkbook.Sheets("原価S直データ")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column

    arrOut = ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol)).value

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub

'============================================
' プロシージャ名 : Transfer_CopyValues_UpDown
' Module        : xl_mod_CostMng1
' 概要          : 基本工事情報業へ小規模工事名と既払高：経費に値を転写
' 引数          : arr1 - 2次元配列（タイトル行＋データ）
'============================================
Public Sub Transfer_CopyValues_UpDown(arr1 As Variant)
    Dim titleRow As Long: titleRow = 1
    Dim rowCount As Long: rowCount = UBound(arr1, 1)
    Dim colNo As Long, colCopyFrom As Long, colCopyTo As Long
    Dim r As Long

    '--- № = 1 の行 → 小規模工事名 を1行下へ転写
    colNo = GetTitleColumn(arr1, "№")
    colCopyFrom = GetTitleColumn(arr1, "小規模工事名")
    colCopyTo = GetTitleColumn(arr1, "小規模工事名")
    If colNo > 0 And colCopyFrom > 0 And colCopyTo > 0 Then
        For r = 2 To rowCount - 1
            If Trim(arr1(r, colNo) & "") = "1" Then
                arr1(r + 1, colCopyTo) = arr1(r, colCopyFrom)
            End If
        Next r
    End If

    '--- № = 3 の行 → 既払高：総額 を1行上へ転写
    colCopyFrom = GetTitleColumn(arr1, "既払高：総額")
    colCopyTo = GetTitleColumn(arr1, "既払高：経費")
    If colNo > 0 And colCopyFrom > 0 And colCopyTo > 0 Then
        For r = 3 To rowCount
            If Trim(arr1(r, colNo) & "") = "3" Then
                arr1(r - 1, colCopyTo) = arr1(r, colCopyFrom)
            End If
        Next r
    End If
End Sub

'============================================
' プロシージャ名 : Transfer_ProcessedConstructionName_Updated
' Module        : xl_mod_CostMng1
' 概要          : 「№」が"1"のたびに「小規模工事名」から分割し、「基本工事コード」「基本工事名」を転写
' 引数          : arr1 - 2次元配列
'============================================
Public Sub Transfer_ProcessedConstructionName_Updated(arr1 As Variant)
    Dim rowCount As Long: rowCount = UBound(arr1, 1)
    Dim colNo As Long, colSmallName As Long, colOutCode As Long, colOutName As Long
    Dim r As Long
    Dim extractedCode As String, extractedName As String, targetValue As String
    Dim isNewGroup As Boolean: isNewGroup = True

    colNo = GetTitleColumn(arr1, "№")
    colSmallName = GetTitleColumn(arr1, "小規模工事名")
    colOutCode = GetTitleColumn(arr1, "基本工事コード")
    colOutName = GetTitleColumn(arr1, "基本工事名")

    If colNo = 0 Or colSmallName = 0 Or colOutCode = 0 Or colOutName = 0 Then
        MsgBox "必要な列が見つからないにゃ", vbExclamation
        Exit Sub
    End If

    For r = 2 To rowCount
        Dim currentNo As String
        currentNo = Trim(arr1(r, colNo) & "")

        If currentNo = "" Then
            isNewGroup = True
        Else
            ' 「№」が"1" または 空白の後であれば、新しいグループとみなす
            If currentNo = "1" Or isNewGroup Then
                targetValue = Trim(arr1(r, colSmallName) & "")
                
                If Mid(targetValue, 10, 1) = " " Or Mid(targetValue, 10, 1) = "" Then
                    extractedCode = Left(targetValue, 9)
                    extractedName = Mid(targetValue, 11)
                Else
                    extractedCode = ""
                    extractedName = ""
                End If

                isNewGroup = False
            End If

            ' データある限り、コードと名前を転写
            arr1(r, colOutCode) = extractedCode
            arr1(r, colOutName) = extractedName
        End If
    Next r
End Sub

'============================================
' プロシージャ名 : Transfer_ConstructionCodeAndName_Final
' Module        : xl_mod_CostMng1
' 概要          : 「枝番工事コード」が"1"のたび1行上の「小規模工事名」を加工し、出力行に転写
' 引数          : arr1 - 2次元配列
'============================================
Public Sub Transfer_ConstructionCodeAndName_Final(arr1 As Variant)
    Dim rowCount As Long: rowCount = UBound(arr1, 1)
    Dim colBranchCode As Long, colSmallName As Long, colNo As Long, colOutCode As Long, colOutName As Long
    Dim r As Long
    Dim sourceValue As String, extractedCode As String, extractedName As String
    Dim isNewGroup As Boolean: isNewGroup = True

    colBranchCode = GetTitleColumn(arr1, "枝番工事コード")
    colSmallName = GetTitleColumn(arr1, "小規模工事名")
    colNo = GetTitleColumn(arr1, "№")
    colOutCode = GetTitleColumn(arr1, "工事コード")
    colOutName = GetTitleColumn(arr1, "工事名")

    If colBranchCode = 0 Or colSmallName = 0 Or colNo = 0 Or colOutCode = 0 Or colOutName = 0 Then
        MsgBox "必要な列が見つからないにゃ", vbExclamation
        Exit Sub
    End If

    For r = 3 To rowCount
        Dim currentVal As String
        currentVal = Trim(arr1(r, colBranchCode) & "")

        ' 空白かどうかチェック
        If currentVal = "" Then
            isNewGroup = True
        Else
            ' 値があり、かつ空白の直後ならリセットして抽出
            If currentVal = "1" Or isNewGroup Then
                sourceValue = Trim(arr1(r - 1, colSmallName) & "")
                If Mid(sourceValue, 10, 1) = " " Or Mid(sourceValue, 10, 1) = "" Then
                    extractedCode = Left(sourceValue, 9)
                    extractedName = Mid(sourceValue, 11)
                Else
                    extractedCode = ""
                    extractedName = ""
                End If
                isNewGroup = False
            End If

            ' 値がある限り転写を継続
            If Trim(arr1(r, colNo) & "") <> "" Then
                arr1(r, colOutCode) = extractedCode
                arr1(r, colOutName) = extractedName
            End If
        End If
    Next r
End Sub


'============================================
' プロシージャ名 : Count_ResetOnBlank_ThroughAll
' Module        : xl_mod_CostMng1
' 概要          : 枝番工事コード列において、カウント値を記入
' 引数          : arr1 - 2次元配列
'============================================
Public Sub Count_ResetOnBlank_ThroughAll(arr1 As Variant)
    Dim rowCount As Long: rowCount = UBound(arr1, 1)
    Dim colTarget As Long
    Dim r As Long
    Dim countVal As Long
    Dim isNewGroup As Boolean: isNewGroup = True

    colTarget = GetTitleColumn(arr1, "枝番工事コード")
    If colTarget = 0 Then
        MsgBox "対象列「枝番工事コード」が見つからないにゃ", vbExclamation
        Exit Sub
    End If

    For r = 2 To rowCount
        If Trim(arr1(r, colTarget) & "") = "" Then
            ' 空白行 → 次が新グループとみなす
            isNewGroup = True
        Else
            If isNewGroup Then
                countVal = 1
                isNewGroup = False
            Else
                countVal = countVal + 1
            End If

            arr1(r, colTarget) = countVal
        End If
    Next r
End Sub



'============================================
' プロシージャ名 : Transfer_AdditionalWorkName
' Module        : xl_mod_CostMng1
' 概要          : 枝番工事コードが空白でない行へ「小規模工事名」を「追加工事名称」に転写
' 引数          : arr1 - 2次元配列
'============================================
Public Sub Transfer_AdditionalWorkName(arr1 As Variant)
    Dim rowCount As Long: rowCount = UBound(arr1, 1)
    Dim colBranchCode As Long, colFrom As Long, colTo As Long, r As Long

    colBranchCode = GetTitleColumn(arr1, "枝番工事コード")
    colFrom = GetTitleColumn(arr1, "小規模工事名")
    colTo = GetTitleColumn(arr1, "追加工事名称")

    If colBranchCode = 0 Or colFrom = 0 Or colTo = 0 Then
        MsgBox "必要な列が見つからないにゃ", vbExclamation
        Exit Sub
    End If

    For r = 2 To rowCount
        If Trim(arr1(r, colBranchCode) & "") <> "" Then
            arr1(r, colTo) = arr1(r, colFrom)
        End If
    Next r
End Sub

'============================================
' プロシージャ名 : RemoveRowsByEmptyCode_Final
' 概要          : 行を削除した後、最後に「№」列を「管理番号」へ変換する
'============================================
Public Sub RemoveRowsByEmptyCode_Final(ByRef arr1 As Variant)
    Dim colBaseCode As Long, colBranchCode As Long, colNo As Long
    Dim r As Long, c As Long, countKeep As Long, rowCount As Long, colCount As Long
    Dim tempArr() As Variant

    colBaseCode = GetTitleColumn(arr1, "基本工事コード")
    colBranchCode = GetTitleColumn(arr1, "枝番工事コード")
    colNo = GetTitleColumn(arr1, "№") ' ここではまだ「№」という名前

    If colBaseCode = 0 Or colBranchCode = 0 Or colNo = 0 Then
        MsgBox "必要な列が見つからないにゃ", vbExclamation
        Exit Sub
    End If

    rowCount = UBound(arr1, 1)
    colCount = UBound(arr1, 2)
    countKeep = 1 ' タイトル行分

    ' --- 1. まず残す行を正しく判定（№=2を基準にする） ---
    For r = 2 To rowCount
        [cite_start] ' №が2の行、またはコードが埋まっている行を保持
        If Trim(arr1(r, colNo) & "") = "2" Then
            countKeep = countKeep + 1
        ElseIf Not (Trim(arr1(r, colBaseCode) & "") = "" Or Trim(arr1(r, colBranchCode) & "") = "") Then
            countKeep = countKeep + 1
        End If
    Next r

    ' --- 2. 新しい配列にコピー ---
    ReDim tempArr(1 To countKeep, 1 To colCount)
    Dim newRow As Long: newRow = 1
    For c = 1 To colCount: tempArr(newRow, c) = arr1(1, c): Next c

    For r = 2 To rowCount
        If Trim(arr1(r, colNo) & "") = "2" Or _
           Not (Trim(arr1(r, colBaseCode) & "") = "" Or Trim(arr1(r, colBranchCode) & "") = "") Then
            newRow = newRow + 1
            For c = 1 To colCount: tempArr(newRow, c) = arr1(r, c): Next c
        End If
    Next r

    ' --- 3. 【重要】最後に「№」列を「管理番号」に変換する ---
    tempArr(1, colNo) = "管理番号"
    For r = 2 To UBound(tempArr, 1)
        Dim fullBranchCode As String: fullBranchCode = Trim(tempArr(r, colBranchCode) & "")
        If fullBranchCode <> "" Then
            ' ハイフンの後ろの数字を抽出して管理番号にセット
            If InStr(fullBranchCode, "-") > 0 Then
                tempArr(r, colNo) = Mid(fullBranchCode, InStrRev(fullBranchCode, "-") + 1)
            Else
                tempArr(r, colNo) = fullBranchCode
            End If
        End If
    Next r

    arr1 = tempArr
End Sub


'============================================
' プロシージャ名 : ClassifyRowsByConstructionCode
' Module        : xl_mod_CostMng1
' 概要          : 「工事コード」がある行には「基本工事」、空白には「枝番工事」と「行分類」列に記入
' 引数          : arr1 - 2次元配列
'============================================
Public Sub ClassifyRowsByConstructionCode(ByRef arr1 As Variant)
    Dim rowCount As Long: rowCount = UBound(arr1, 1)
    Dim colProjectCode As Long, colClassify As Long, r As Long

    colProjectCode = GetTitleColumn(arr1, "工事コード")
    colClassify = GetTitleColumn(arr1, "行分類")

    If colProjectCode = 0 Or colClassify = 0 Then
        MsgBox "必要な列が見つからないにゃ", vbExclamation
        Exit Sub
    End If

    For r = 2 To rowCount
        If Trim(arr1(r, colProjectCode) & "") <> "" Then
            arr1(r, colClassify) = "枝番工事"
        Else
            arr1(r, colClassify) = "基本工事"
        End If
    Next r
End Sub

'============================================
' プロシージャ名 : OutputArrayToTable_Overwrite
' Module        : xl_mod_CostMng1
' 概要          : 配列arr1の1行目（タイトル）と一致する列のみテーブル t_原価S_temp へ上書き出力
' 引数          : arr1 - 2次元配列
'============================================
Public Sub OutputArrayToTable_Overwrite(arr1 As Variant)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim headerRowInSheet As Long: headerRowInSheet = 6
    Dim colStart As Long: colStart = 2
    Dim lastCol As Long, r As Long, c As Long, arrRowCount As Long, tblRowCount As Long
    Dim arrCol As Long, sheetCol As Long, titleName As String
    Dim dictColMap As Object

    Set ws = ThisWorkbook.Sheets("原価S_temp")
    On Error Resume Next
    Set tbl = ws.ListObjects("t_原価S_temp")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "テーブル t_原価S_temp が見つからないにゃ", vbExclamation
        Exit Sub
    End If

    arrRowCount = UBound(arr1, 1) - 1
    tblRowCount = tbl.ListRows.Count

    Do While tbl.ListRows.Count < arrRowCount
        tbl.ListRows.Add
    Loop
    Do While tbl.ListRows.Count > arrRowCount
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop

    Set dictColMap = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(headerRowInSheet, ws.Columns.Count).End(xlToLeft).Column
    For arrCol = LBound(arr1, 2) To UBound(arr1, 2)
        titleName = Trim(arr1(1, arrCol) & "")
        For sheetCol = colStart To lastCol
            If Trim(ws.Cells(headerRowInSheet, sheetCol).value & "") = titleName Then
                dictColMap(arrCol) = sheetCol
                Exit For
            End If
        Next sheetCol
    Next arrCol

    For r = 2 To UBound(arr1, 1)
        For arrCol = LBound(arr1, 2) To UBound(arr1, 2)
            If dictColMap.Exists(arrCol) Then
                ws.Cells(r + 5, dictColMap(arrCol)).value = arr1(r, arrCol)
            End If
        Next arrCol
    Next r
End Sub

'============================================
' 関数名 : GetTitleColumn
' Module : xl_mod_CostMng1
' 概要   : 指定されたタイトル名に対応する列番号を返す
' 引数   : arr - 配列, title - タイトル名（列名）
' 戻り値 : 列番号（Long）※見つからない場合は 0
'============================================
Public Function GetTitleColumn(arr As Variant, title As String) As Long
    Dim j As Long
    For j = LBound(arr, 2) To UBound(arr, 2)
        If Trim(arr(1, j) & "") = Trim(title) Then
            GetTitleColumn = j
            Exit Function
        End If
    Next j
    GetTitleColumn = 0
End Function

