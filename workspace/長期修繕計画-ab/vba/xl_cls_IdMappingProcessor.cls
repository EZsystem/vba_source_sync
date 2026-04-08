VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_IdMappingProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------
' Class Module: cls_IdMappingProcessor
' 説明　：she02_IdMappingSingle/Doubleのロジックを完全再現した高速化クラス
' 復元　：Trimなし、区切り文字なし、種類列あり、Idpaは現在行→上行へコピー
'-------------------------------------
Option Explicit

' --- メンバ変数 ---
Private p_wsBreakdown As Worksheet
Private p_wsCategory As Worksheet
Private p_tblBreakdown As ListObject
Private p_tblCategory As ListObject

'============================================
' 初期化
'============================================
Private Sub Class_Initialize()
    On Error Resume Next
    Set p_wsBreakdown = ThisWorkbook.Worksheets("内訳")
    Set p_wsCategory = ThisWorkbook.Worksheets("分類")
    
    If Not p_wsBreakdown Is Nothing Then
        Set p_tblBreakdown = p_wsBreakdown.ListObjects("tbl_内訳")
    End If
    
    If Not p_wsCategory Is Nothing Then
        Set p_tblCategory = p_wsCategory.ListObjects("tbl_内訳ID")
    End If
    On Error GoTo 0
End Sub

'============================================
' メイン実行メソッド
'============================================
Public Sub ExecuteFullProcess()
    
    If p_tblBreakdown Is Nothing Or p_tblCategory Is Nothing Then
        MsgBox "必要なテーブル（tbl_内訳, tbl_内訳ID）が見つかりません。", vbCritical
        Exit Sub
    End If

    ' --- 高速化設定 ---
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    On Error GoTo ErrorHandler
    
    ' 1. マッピング実行（オリジナルロジック完全再現）
    Call ProcessMapping_Original
    
    ' 2. Idpa補正実行（オリジナルロジック完全再現）
    Call ProcessIdpa_Original
    
    ' --- 完了メッセージ ---
    ' 呼び出し元で表示するためここでは省略可ですが、念のため
    ' MsgBox "処理が完了しました。", vbInformation
    
Finally:
    ' --- 設定復元 ---
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Finally
    
End Sub

'============================================
' 1. マッピング処理
' ロジック：5列連結、Trimなし、区切り文字なし
'============================================
Private Sub ProcessMapping_Original()
    Dim dict As Object
    Dim dataArr As Variant, catArr As Variant
    Dim resultArr() As Variant
    Dim r As Long
    Dim keyStr As String
    
    ' --- 辞書作成 ---
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 分類マスタ読み込み
    If p_tblCategory.DataBodyRange Is Nothing Then Exit Sub
    catArr = p_tblCategory.DataBodyRange.value
    
    ' 列位置取得
    Dim cL As Long, cM As Long, cT As Long, cC As Long, cU As Long, cID As Long
    cL = GetColIdx(p_tblCategory, "大分類")
    cM = GetColIdx(p_tblCategory, "中分類")
    cT = GetColIdx(p_tblCategory, "種類")       ' ★種類列（必須）
    cC = GetColIdx(p_tblCategory, "周期")
    cU = GetColIdx(p_tblCategory, "更新周期")
    cID = GetColIdx(p_tblCategory, "内訳ID")
    
    For r = 1 To UBound(catArr, 1)
        ' ★she02完全再現：CStrのみ、Trimなし、単純連結
        keyStr = CStr(catArr(r, cL)) & CStr(catArr(r, cM)) & _
                 CStr(catArr(r, cT)) & CStr(catArr(r, cC)) & _
                 CStr(catArr(r, cU))
        
        If Not dict.Exists(keyStr) Then
            dict.Add keyStr, CStr(catArr(r, cID))
        End If
    Next r
    
    ' --- 内訳テーブル適用 ---
    If p_tblBreakdown.DataBodyRange Is Nothing Then Exit Sub
    dataArr = p_tblBreakdown.DataBodyRange.value
    
    ' 列位置取得
    Dim dL As Long, dM As Long, dT As Long, dC As Long, dU As Long, dID As Long
    dL = GetColIdx(p_tblBreakdown, "大分類")
    dM = GetColIdx(p_tblBreakdown, "中分類")
    dT = GetColIdx(p_tblBreakdown, "種類")       ' ★種類列（必須）
    dC = GetColIdx(p_tblBreakdown, "周期")
    dU = GetColIdx(p_tblBreakdown, "更新周期")
    dID = GetColIdx(p_tblBreakdown, "内訳ID")
    
    ' 結果用配列確保
    ReDim resultArr(1 To UBound(dataArr, 1), 1 To 1)
    
    For r = 1 To UBound(dataArr, 1)
        ' ★she02完全再現：CStrのみ、Trimなし、単純連結
        keyStr = CStr(dataArr(r, dL)) & CStr(dataArr(r, dM)) & _
                 CStr(dataArr(r, dT)) & CStr(dataArr(r, dC)) & _
                 CStr(dataArr(r, dU))
                 
        If dict.Exists(keyStr) Then
            resultArr(r, 1) = dict(keyStr)
        Else
            resultArr(r, 1) = "該当なし"
        End If
    Next r
    
    ' 一括書き込み
    p_tblBreakdown.ListColumns(dID).DataBodyRange.value = resultArr
    
End Sub

'============================================
' 2. Idpa補正処理
' ロジック：Idpa="b" (Trimなし) なら 現在行(i)のIDを 上行(i-1)へコピー
'============================================
Private Sub ProcessIdpa_Original()
    Dim dataArr As Variant
    Dim idArr As Variant
    Dim r As Long
    Dim colIdpa As Long, colID As Long
    Dim isUpdated As Boolean
    
    colIdpa = GetColIdx(p_tblBreakdown, "Idpa")
    colID = GetColIdx(p_tblBreakdown, "内訳ID")
    
    If colIdpa = 0 Or colID = 0 Then Exit Sub
    If p_tblBreakdown.DataBodyRange Is Nothing Then Exit Sub
    
    ' 判定用データと書き込み用データ配列
    dataArr = p_tblBreakdown.DataBodyRange.value
    idArr = p_tblBreakdown.ListColumns(colID).DataBodyRange.value
    
    ' 2行目からループ
    For r = 2 To UBound(dataArr, 1)
        ' ★she02完全再現：CStrのみ、Trimなし、"b"判定
        If CStr(dataArr(r, colIdpa)) = "b" Then
            ' ★she02完全再現：現在行(r) のIDを 1行上(r-1) にコピー
            idArr(r - 1, 1) = idArr(r, 1)
            isUpdated = True
        End If
    Next r
    
    ' 変更があった場合のみ書き込み
    If isUpdated Then
        p_tblBreakdown.ListColumns(colID).DataBodyRange.value = idArr
    End If
    
End Sub

' ヘルパー関数
Private Function GetColIdx(tbl As ListObject, colName As String) As Long
    On Error Resume Next
    GetColIdx = tbl.ListColumns(colName).Index
    On Error GoTo 0
End Function

