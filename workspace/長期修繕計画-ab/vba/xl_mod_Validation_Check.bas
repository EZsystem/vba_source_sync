Attribute VB_Name = "xl_mod_Validation_Check"

'-------------------------------------
' Module: xl_mod_Validation_Check
' 説明  ：内訳テーブルにおけるO-IDの重複チェック（グループ全着色版）
' 構成  ：初期化 -> 検証 -> 実行 -> 終了処理 [cite: 11]
'-------------------------------------
Option Explicit

' マジックナンバーの定義（ルール 15） [cite: 15]
Private Const COLOR_YELLOW As Long = 65535 ' vbYellow
Private Const DEST_TABLE_NAME As String = "xt_内訳" ' ルール 6-1 [cite: 9]
Private Const TARGET_COL_A As String = "A指示20260310"
Private Const TARGET_COL_ID As String = "O-ID"

''' <summary>
''' 重複するO-IDグループを抽出し、グループ内にA指示の値が1つでもある場合は全対象行を黄色く塗る
''' </summary>
Public Sub Highlight_DuplicateOIDs_All()
    Dim shInner As Worksheet: Set shInner = ThisWorkbook.Worksheets("内訳")
    Dim xtInner As ListObject
    
    ' 1. 初期化・検証 [cite: 11]
    On Error Resume Next
    Set xtInner = shInner.ListObjects(DEST_TABLE_NAME)
    If xtInner Is Nothing Then Set xtInner = shInner.ListObjects("tbl_内訳")
    On Error GoTo 0
    
    If xtInner Is Nothing Then
        MsgBox "テーブルが見つかりません。", vbCritical
        Exit Sub
    End If

    ' 列インデックスの取得
    Dim colOID As Long: colOID = Get_ColumnIndex_Robust(xtInner, TARGET_COL_ID)
    Dim colA As Long: colA = Get_ColumnIndex_Robust(xtInner, TARGET_COL_A)
    
    If colOID = 0 Or colA = 0 Then
        MsgBox "必要な列が見つかりません。", vbCritical
        Exit Sub
    End If

    ' 2. 実行：高速モード開始 [cite: 20]
    Call Fast_Mode_Toggle(True)

    ' 前処理：色塗りを解除 [cite: 11]
    If Not xtInner.DataBodyRange Is Nothing Then
        xtInner.ListColumns(colOID).DataBodyRange.Interior.ColorIndex = xlNone
    Else
        GoTo Finalize
    End If

    ' 辞書の準備
    Dim dictCount As Object: Set dictCount = CreateObject("Scripting.Dictionary")
    Dim dictHasAValue As Object: Set dictHasAValue = CreateObject("Scripting.Dictionary")
    Dim dataArr As Variant: dataArr = xtInner.DataBodyRange.value
    Dim i As Long, oidValue As String
    
    ' --- 3. パス1：統計情報の収集 ---
    ' テーブル全体を走査し、IDごとの出現数と「A指示」の値有無を記録
    For i = 1 To UBound(dataArr, 1)
        oidValue = Normalize_Text(CStr(dataArr(i, colOID)))
        If oidValue <> "" Then
            ' 重複件数をカウント
            dictCount(oidValue) = dictCount(oidValue) + 1
            ' グループ内に「A指示」の値があるかチェック
            If Len(Trim(CStr(dataArr(i, colA)))) > 0 Then
                dictHasAValue(oidValue) = True
            End If
        End If
    Next i

    ' --- 4. パス2：条件判定とグループ一括着色 ---
    ' 条件：(出現数が1より大きい) かつ (そのIDグループの誰かがA指示に値を持っている)
    For i = 1 To UBound(dataArr, 1)
        oidValue = Normalize_Text(CStr(dataArr(i, colOID)))
        
        If dictCount.Exists(oidValue) Then
            ' 重複しているグループであり、かつ条件1を満たすグループに属している場合
            If dictCount(oidValue) > 1 And dictHasAValue.Exists(oidValue) Then
                ' 行ごとのA指示の値有無に関わらず、グループ全員を着色
                xtInner.DataBodyRange.Cells(i, colOID).Interior.Color = COLOR_YELLOW
            End If
        End If
    Next i

Finalize:
    ' 5. 終了処理 [cite: 11]
    Call Fast_Mode_Toggle(False)
    Call Notify_Smart_Popup("重複グループの一括着色が完了しました。", "完了") '[cite: 20]

End Sub
' ← プロシージャの終わり [cite: 12]

