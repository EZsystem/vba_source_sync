VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "com_clsArrayHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: com_clsArrayHelper
' 説明：2次元配列（主にExcel取込データ）の解析、列探索、一括変換を行う。
'       acc_clsDataCleaner と連携し、メモリ上での高速クレンジングを実現。
' 更新日：2026/03/24（MyLauncher統合・データクレンジング連携版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pData        As Variant
Private pCleaner     As acc_clsDataCleaner
Private pLBound1     As Long
Private pUBound1     As Long
Private pLBound2     As Long
Private pUBound2     As Long

'===========================================================
' サブルーチン名 : Init
' 引数 : sourceArray - 処理対象の2次元配列
' 機能 : 配列をセットし、解析の準備（境界値の取得）を行う。
'===========================================================
Public Sub Init(ByRef sourceArray As Variant)
    pData = sourceArray
    
    If IsEmpty(pData) Then Exit Sub
    
    ' 2次元配列であることを確認して境界を保持
    On Error Resume Next
    pLBound1 = LBound(pData, 1): pUBound1 = UBound(pData, 1)
    pLBound2 = LBound(pData, 2): pUBound2 = UBound(pData, 2)
    If Err.Number <> 0 Then
        Err.Raise 1005, "com_clsArrayHelper.Init", "2次元配列を指定してください。"
    End If
    On Error GoTo 0
    
    ' クリーナーの準備
    Set pCleaner = New acc_clsDataCleaner
    pCleaner.Init
End Sub

' --- プロパティ ---
Public Property Get Data() As Variant: Data = pData: End Property
Public Property Get RowCount() As Long: RowCount = pUBound1 - pLBound1 + 1: End Property
Public Property Get ColCount() As Long: ColCount = pUBound2 - pLBound2 + 1: End Property

'===========================================================
' 関数名 : GetColIndex
' 引数 : colName - 探索する列名（配列の1行目をヘッダーと見なす）
' 戻り値 : 列インデックス（見つからない場合は -1）
'===========================================================
Public Function GetColIndex(ByVal colName As String) As Long
    GetColIndex = -1
    If IsEmpty(pData) Then Exit Function
    
    Dim j As Long
    Dim target As String: target = UCase(Trim(colName))
    
    ' 1行目をループして列名を探す
    For j = pLBound2 To pUBound2
        If UCase(Trim(CStr(pData(pLBound1, j)))) = target Then
            GetColIndex = j
            Exit Function
        End If
    Next j
End Function

'===========================================================
' 関数名 : GetValue
' 引数 : row, colIndex
' 戻り値 : Variant
'===========================================================
Public Function GetValue(ByVal row As Long, ByVal colIndex As Long) As Variant
    If row < pLBound1 Or row > pUBound1 Or colIndex < pLBound2 Or colIndex > pUBound2 Then
        GetValue = Empty
    Else
        GetValue = pData(row, colIndex)
    End If
End Function

'===========================================================
' 関数名 : ToCleanedArray
' 戻り値 : クレンジング済みの2次元配列
' 機能 : 配列内の全要素に対し、Trim・改行除去・Null置換を一括実行する。
'===========================================================
Public Function ToCleanedArray() As Variant
    If IsEmpty(pData) Then Exit Function
    
    Dim result As Variant
    result = pData
    
    Dim i As Long, j As Long
    For i = pLBound1 To pUBound1
        For j = pLBound2 To pUBound2
            ' DataCleanerを使用して全要素を安全な文字列に変換
            result(i, j) = pCleaner.CleanText(pData(i, j))
        Next j
    Next j
    
    ToCleanedArray = result
End Function

'===========================================================
' 関数名 : ExtractColumn
' 引数 : colIndex - 抽出したい列番号
' 戻り値 : 1次元配列
'===========================================================
Public Function ExtractColumn(ByVal colIndex As Long) As Variant
    If colIndex < pLBound2 Or colIndex > pUBound2 Then Exit Function
    
    Dim arr() As Variant
    ReDim arr(pLBound1 To pUBound1)
    
    Dim i As Long
    For i = pLBound1 To pUBound1
        arr(i) = pData(i, colIndex)
    Next i
    
    ExtractColumn = arr
End Function

