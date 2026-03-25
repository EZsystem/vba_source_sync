VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xl_clsRangeAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: xl_clsRangeAccessor
' 説明：Excelのレンジ操作（配列変換、最終行取得、書き込み）を管理する。
'       参照設定に依存しないレイトバインディング方式を採用。
' 更新日：2026/03/24（MyLauncher統合・堅牢レンジ操作版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pErr As com_clsErrorUtility

'===========================================================
' サブルーチン名 : Init
' 引数 : errorUtil - com_clsErrorUtility インスタンス
' 機能 : エラーハンドラを初期化する。
'===========================================================
Public Sub Init(ByRef errorUtil As com_clsErrorUtility)
    Set pErr = errorUtil
End Sub

'===========================================================
' 関数名 : RangeToArray
' 引数 : rng - Excel.Range オブジェクト (Object)
' 戻り値 : 2次元配列 (Variant)
' 機能 : 指定されたセル範囲の値を一括で配列として取得する。
'===========================================================
Public Function RangeToArray(ByVal rng As Object) As Variant
    On Error GoTo Err_Handler
    If rng Is Nothing Then Exit Function
    
    ' セルが1つだけの場合、Valueをそのまま返すと配列にならないため補正
    If rng.Cells.count = 1 Then
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = rng.Value
        RangeToArray = arr
    Else
        RangeToArray = rng.Value
    End If
    Exit Function

Err_Handler:
    pErr.Notify_Smart_Popup "範囲の取得に失敗しました: " & Err.description, "Range Error", vbCritical
    RangeToArray = Empty
End Function

'===========================================================
' サブルーチン名 : ArrayToRange
' 引数 : targetRng - 書き込み先の左上セル、または範囲 (Object)
'        data - 書き込む2次元配列
' 機能 : 配列データをセル範囲に一括で書き込む。
'===========================================================
Public Sub ArrayToRange(ByVal targetRng As Object, ByVal Data As Variant)
    On Error GoTo Err_Handler
    If targetRng Is Nothing Or IsEmpty(Data) Then Exit Sub
    
    ' 配列のサイズに合わせてリサイズして一括書き込み
    Dim r As Long, c As Long
    r = UBound(Data, 1) - LBound(Data, 1) + 1
    c = UBound(Data, 2) - LBound(Data, 2) + 1
    
    targetRng.Resize(r, c).Value = Data
    Exit Sub

Err_Handler:
    pErr.Notify_Smart_Popup "データの書き込みに失敗しました: " & Err.description, "Write Error", vbCritical
End Sub

'===========================================================
' 関数名 : GetLastRow
' 引数 : ws - 対象シート (Object)
'        col - チェックする列番号 (Default: 1)
' 戻り値 : Long (最終行番号)
'===========================================================
Public Function GetLastRow(ByVal ws As Object, Optional ByVal col As Long = 1) As Long
    On Error GoTo Err_Handler
    GetLastRow = ws.Cells(ws.Rows.count, col).End(-4162).row ' -4162 = xlUp
    Exit Function

Err_Handler:
    GetLastRow = 0
End Function

'===========================================================
' 関数名 : GetLastColumn
' 引数 : ws - 対象シート (Object)
'        row - チェックする行番号 (Default: 1)
' 戻り値 : Long (最終列番号)
'===========================================================
Public Function GetLastColumn(ByVal ws As Object, Optional ByVal row As Long = 1) As Long
    On Error GoTo Err_Handler
    GetLastColumn = ws.Cells(row, ws.Columns.count).End(-4159).Column ' -4159 = xlToLeft
    Exit Function

Err_Handler:
    GetLastColumn = 0
End Function
