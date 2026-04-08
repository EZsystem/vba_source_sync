Attribute VB_Name = "xl_mod_DataTransfer"

'-------------------------------------
' Module: xl_mod_DataTransfer
' 説明  ：複数シートから黄色セルを検出し、ID(A列)+内訳ID(C7)で照合。
'         値は「A指示20260310」列のみに記入し、一致タイトル列は着色のみ行う。
' 構成  ：初期化 -> 検証 -> 実行 -> 終了処理 '[cite: 11]
'-------------------------------------
Option Explicit ' 必須宣言 '[cite: 9]

' マジックナンバーの定義（ルール 15） '[cite: 15]
Private Const START_ROW As Long = 7    ' データ開始行
Private Const HEADER_ROW As Long = 6   ' 見出し行
Private Const SRC_REF_CELL As String = "C7" ' コピー元の照合キーセル
Private Const COLOR_YELLOW As Long = 65535 ' vbYellow
Private Const SRC_COL_START As Long = 2   ' B列
Private Const SRC_COL_END As Long = 13    ' M列
Private Const DEST_TABLE_NAME As String = "xt_内訳" ' Excelテーブル接頭辞 '[cite: 2, 13]
Private Const FIXED_TARGET_COL As String = "A指示20260310" ' 唯一の転記先列

''' <summary>
''' 黄色セルの値を固定列へ転記し、対応するタイトル列を識別用に塗りつぶす '[cite: 6]
''' </summary>
Public Sub Execute_YellowCellTransfer()
    Dim shDest As Worksheet: Set shDest = ThisWorkbook.Worksheets("内訳") '[cite: 10]
    Dim xtInner As ListObject
    
    ' 1. 検証（バリデーション） '[cite: 11]
    On Error Resume Next
    Set xtInner = shDest.ListObjects(DEST_TABLE_NAME)
    If xtInner Is Nothing Then Set xtInner = shDest.ListObjects("tbl_内訳")
    On Error GoTo 0
    
    If xtInner Is Nothing Then
        MsgBox "出力先テーブルが見つかりません。", vbCritical
        Exit Sub
    End If

    ' 列インデックスの堅牢な取得
    Dim colOID As Long: colOID = Get_ColumnIndex_Robust(xtInner, "O-ID")
    Dim colUchiID As Long: colUchiID = Get_ColumnIndex_Robust(xtInner, "内訳ID")
    Dim colFixed As Long: colFixed = Get_ColumnIndex_Robust(xtInner, FIXED_TARGET_COL)
    
    If colOID = 0 Or colUchiID = 0 Or colFixed = 0 Then
        MsgBox "必要な列（O-ID, 内訳ID, または " & FIXED_TARGET_COL & "）が見つかりません。", vbCritical
        Exit Sub
    End If

    ' シート選択状態の検証
    Dim selectedSheets As Object: Set selectedSheets = ActiveWindow.selectedSheets
    Dim hasValidSheet As Boolean: hasValidSheet = False
    Dim sh As Object
    For Each sh In selectedSheets
        If sh.Name <> shDest.Name And TypeOf sh Is Worksheet Then
            hasValidSheet = True
            Exit For
        End If
    Next sh

    If Not hasValidSheet Then
        MsgBox "処理対象のコピー元シートを選択した状態で実行してください。", vbExclamation
        Exit Sub
    End If

    ' 2. 実行：高速モード開始 '[cite: 11, 20]
    Call Fast_Mode_Toggle(True)

    ' --- 前処理：データクリアおよび色塗り解除 ---
    If Not xtInner.DataBodyRange Is Nothing Then
        ' 固定転記先列の値のみをクリア
        xtInner.ListColumns(colFixed).DataBodyRange.ClearContents
        ' テーブル全域の色塗りをリセット '[cite: 15]
        xtInner.DataBodyRange.Interior.ColorIndex = xlNone
    End If

    ' 複合キーによる照合用辞書の作成
    Dim dictMatch As Object: Set dictMatch = CreateObject("Scripting.Dictionary")
    Dim destData As Variant: destData = xtInner.DataBodyRange.value
    Dim i As Long
    For i = 1 To UBound(destData, 1)
        Dim compositeKey As String
        compositeKey = Normalize_Text(CStr(destData(i, colOID))) & "|" & _
                       Normalize_Text(CStr(destData(i, colUchiID)))
        If compositeKey <> "|" Then dictMatch(compositeKey) = i
    Next i

    ' 3. 各シートの転記・着色ロジック
    Dim srcWs As Worksheet
    Dim r As Long, c As Long, varA As String, varB As Variant
    Dim valC7 As String, searchKey As String
    Dim srcHeader As String, colDynamic As Long
    
    For Each sh In selectedSheets
        If sh.Name <> shDest.Name And TypeOf sh Is Worksheet Then
            Set srcWs = sh
            valC7 = Normalize_Text(CStr(srcWs.Range(SRC_REF_CELL).value))
            
            Dim lastRowSource As Long: lastRowSource = srcWs.Cells(srcWs.Rows.count, "B").End(xlUp).Row
            
            If lastRowSource >= START_ROW And valC7 <> "" Then
                For r = START_ROW To lastRowSource
                    varA = Normalize_Text(CStr(srcWs.Cells(r, "A").value))
                    searchKey = varA & "|" & valC7
                    
                    If dictMatch.Exists(searchKey) Then
                        For c = SRC_COL_START To SRC_COL_END
                            If srcWs.Cells(r, c).Interior.Color = COLOR_YELLOW Then
                                varB = srcWs.Cells(r, c).value
                                
                                ' 【修正】値の記入は「固定列」のみに行う
                                xtInner.DataBodyRange.Cells(dictMatch(searchKey), colFixed).value = varB
                                
                                ' 【修正】タイトルが一致する列は「着色」のみ行う
                                srcHeader = CStr(srcWs.Cells(HEADER_ROW, c).value)
                                colDynamic = Get_ColumnIndex_Robust(xtInner, srcHeader)
                                
                                If colDynamic > 0 Then
                                    ' 値は入れず、セルを黄色く塗るのみ
                                    xtInner.DataBodyRange.Cells(dictMatch(searchKey), colDynamic).Interior.Color = COLOR_YELLOW
                                End If
                            End If
                        Next c
                    End If
                Next r
            End If
        End If
    Next sh

    ' 4. 出力・終了処理 '[cite: 11]
    Call Fast_Mode_Toggle(False)
    Call Notify_Smart_Popup("転記（固定列）および識別着色が完了しました。", "完了") '[cite: 19]

End Sub
' ← プロシージャの終わり '[cite: 12]


