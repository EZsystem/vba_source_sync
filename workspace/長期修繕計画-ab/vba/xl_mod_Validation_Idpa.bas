Attribute VB_Name = "xl_mod_Validation_Idpa"

'-------------------------------------
' Module: xl_mod_Validation_Idpa
' 説明  : O-IDとIDnewの奇数・偶数不一致（パリティ不整合）チェックおよびIdpa列の着色
' 構成  : 初期化 -> 検証 -> 実行 -> 終了処理
'-------------------------------------
Option Explicit

' マジックナンバーの定義（ルール 15） [cite: 15]
Private Const COLOR_YELLOW As Long = 65535 ' vbYellow
Private Const DEST_TABLE_NAME As String = "xt_内訳" ' ルール 6-1 [cite: 13]
Private Const COL_NAME_IDPA As String = "Idpa"
Private Const COL_NAME_IDNEW As String = "IDnew"
Private Const COL_NAME_OID As String = "O-ID"

''' <summary>
''' O-IDとIDnewの奇数・偶数が一致していない行のIdpaセルを黄色く塗る
''' </summary>
Public Sub Check_IdpaParityMismatch()
    Dim shInner As Worksheet: Set shInner = ThisWorkbook.Worksheets("内訳") ' ルール 6-2 [cite: 10]
    Dim xtInner As ListObject
    
    ' 1. 初期化・検証（バリデーション）
    On Error Resume Next
    Set xtInner = shInner.ListObjects(DEST_TABLE_NAME)
    If xtInner Is Nothing Then Set xtInner = shInner.ListObjects("tbl_内訳")
    On Error GoTo 0
    
    If xtInner Is Nothing Then
        MsgBox "テーブル '" & DEST_TABLE_NAME & "' が見つかりません。", vbCritical
        Exit Sub
    End If

    ' 共通関数による列番号の堅牢な取得（ルール 18） [cite: 18]
    Dim colIdxIdpa As Long: colIdxIdpa = Get_ColumnIndex_Robust(xtInner, COL_NAME_IDPA)
    Dim colIdxIdnew As Long: colIdxIdnew = Get_ColumnIndex_Robust(xtInner, COL_NAME_IDNEW)
    Dim colIdxOid As Long: colIdxOid = Get_ColumnIndex_Robust(xtInner, COL_NAME_OID)
    
    If colIdxIdpa = 0 Or colIdxIdnew = 0 Or colIdxOid = 0 Then
        MsgBox "必要な列（Idpa, IDnew, O-ID）が見つかりません。", vbCritical
        Exit Sub
    End If

    ' 2. 実行：高速モード開始
    Call Fast_Mode_Toggle(True)

    ' 前処理：Idpa列の着色をクリア（ユーザー要望）
    If Not xtInner.DataBodyRange Is Nothing Then
        xtInner.ListColumns(colIdxIdpa).DataBodyRange.Interior.ColorIndex = xlNone
    Else
        GoTo Finalize
    End If

    ' 3. 行スキャンと条件判定
    Dim dataArr As Variant: dataArr = xtInner.DataBodyRange.value
    Dim i As Long
    Dim valOid As Variant, valIdnew As Variant
    Dim isInvalid As Boolean
    
    For i = 1 To UBound(dataArr, 1)
        isInvalid = False
        valOid = dataArr(i, colIdxOid)
        valIdnew = dataArr(i, colIdxIdnew)
        
        ' 両方の値が数値である場合のみ判定を実施
        If IsNumeric(valOid) And IsNumeric(valIdnew) Then
            ' 条件：奇数・偶数が一致していない（パリティが異なる）場合に対象
            ' (O-ID Mod 2) と (IDnew Mod 2) の結果が異なる場合に不整合とみなす
            If (CLng(valOid) Mod 2) <> (CLng(valIdnew) Mod 2) Then
                isInvalid = True
            End If
        End If
        
        ' 判定結果に基づきIdpaセルを着色
        If isInvalid Then
            xtInner.DataBodyRange.Cells(i, colIdxIdpa).Interior.Color = COLOR_YELLOW
        End If
    Next i

Finalize:
    ' 4. 出力・終了処理
    Call Fast_Mode_Toggle(False)
    Call Notify_Smart_Popup("奇数・偶数の整合性チェックが完了しました。", "完了")
End Sub
' ← プロシージャの終わり [cite: 12]


