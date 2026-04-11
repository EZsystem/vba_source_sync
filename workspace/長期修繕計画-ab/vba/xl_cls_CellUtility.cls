VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_CellUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: cls_CellUtility
' 説明　：セルの値変換・マーク記入などのユーティリティクラス
' 作成日：2025/11/14
'-------------------------------------
Option Explicit

'============================================
' メソッド名 : ConvertToNarrow
' 概要       : 範囲内の文字列を半角に変換（数値・数式は除外）
' 引数       : rng As Range
'============================================
Public Sub ConvertToNarrow(rng As Range)
    Dim cell As Range, area As Range
    Dim val As Variant
    
    On Error Resume Next
    ' 定数セルのみ対象（数式除外）
    Dim targetRange As Range
    Set targetRange = rng.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    
    If targetRange Is Nothing Then Exit Sub
    
    For Each area In targetRange.Areas
        For Each cell In area.Cells
            val = cell.Value2
            If VarType(val) = vbString Then
                ' 全角→半角変換
                Dim newVal As String
                newVal = StrConv(val, vbNarrow)
                ' 必要ならスペース変換など追加
                newVal = Replace(newVal, ChrW(&H3000), " ")
                
                If val <> newVal Then cell.Value2 = newVal
            End If
        Next cell
    Next area
End Sub

'============================================
' メソッド名 : ConvertToWide
' 概要       : 範囲内の値を全角に変換（オプションで数値除外）
' 引数       : rng As Range
'            : skipNumeric As Boolean (Trueなら数値を変換しない)
'============================================
Public Sub ConvertToWide(rng As Range, Optional skipNumeric As Boolean = True)
    Dim cell As Range, area As Range
    Dim val As Variant
    
    On Error Resume Next
    Dim targetRange As Range
    Set targetRange = rng.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    
    If targetRange Is Nothing Then Exit Sub
    
    For Each area In targetRange.Areas
        For Each cell In area.Cells
            val = cell.Value2
            If Not IsError(val) And Not IsEmpty(val) Then
                Dim doConvert As Boolean
                doConvert = True
                
                If skipNumeric Then
                    Select Case VarType(val)
                        Case vbDouble, vbDecimal, vbCurrency, vbDate
                            doConvert = False
                    End Select
                End If
                
                If doConvert Then
                    ' 全角変換（スペース含む）
                    Dim s As String
                    s = StrConv(CStr(val), vbWide)
                    s = Replace(s, " ", "　")
                    cell.value = s
                End If
            End If
        Next cell
    Next area
End Sub

'============================================
' メソッド名 : MarkCircleAbove
' 概要       : "○"があるセルの1つ上に"○"を記入する
' 引数       : rng As Range
' 戻り値     : Long (記入した数)
'============================================
Public Function MarkCircleAbove(rng As Range) As Long
    Dim cell As Range
    Dim count As Long
    
    For Each cell In rng
        If cell.value = "○" Then
            ' 範囲の先頭行より上には書かないチェック
            If cell.Row > 1 Then ' シート最上行チェック
                If cell.Row > rng.Row Or rng.Rows.count = 1 Then ' 範囲内相対チェック不要なら単純にOffset
                    With cell.Offset(-1, 0)
                        .value = "○"
                        count = count + 1
                    End With
                End If
            End If
        End If
    Next cell
    MarkCircleAbove = count
End Function
