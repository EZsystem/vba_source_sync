VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CostMng1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'============================================
' シート: 原価S直データ
' 説明 : シート内で使用するデータ範囲取得関数
' 作成日: 2025/07/01
'============================================

Option Explicit

'============================================
' マクロ登録用: シート内から呼び出し可能な処理
'============================================
Public Sub 実行_データ範囲を表示()

    Dim rng As Range
    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = Me.Cells(Me.Rows.Count, "A").End(xlUp).row
    lastCol = Me.Cells(6, Me.Columns.Count).End(xlToLeft).Column

    Set rng = Me.Range(Me.Cells(7, 1), Me.Cells(lastRow, lastCol))

    rng.Select  ' 範囲選択（オプション）
    MsgBox "対象範囲：" & rng.Address

End Sub

