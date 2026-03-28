Attribute VB_Name = "xl_mod_RangeUtilities"
'-------------------------------------
' Module: xl_mod_RangeUtilities
' 説明  : Range／セル操作の拡張関数
' 作成日: 2025/05/04
' 更新日: -
'-------------------------------------
Option Explicit

'=================================================
' サブルーチン名 : FindAndReplace
' 説明   : 範囲内で検索と置換
'=================================================
Public Sub FindAndReplace(rng As Range, findText As String, replaceText As String)
    rng.Replace what:=findText, Replacement:=replaceText, LookAt:=xlPart
End Sub

'=================================================
' サブルーチン名 : ClearRange
' 説明   : 範囲をクリア
'=================================================
Public Sub ClearRange(rng As Range)
    rng.Clear
End Sub

'=================================================
' サブルーチン名 : AutoFitColumns
' 説明   : 列幅を自動調整
'=================================================
Public Sub AutoFitColumns(rng As Range)
    rng.EntireColumn.AutoFit
End Sub

