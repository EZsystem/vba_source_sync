VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_PrintSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: cls_PrintSetting
' 説明　：印刷範囲設定・フッター設定・ページ数カウントを行うクラス
' 作成日：2025/11/14
'-------------------------------------
Option Explicit

' --- プロパティ（設定値） ---
Private pStartRow As Long           ' データ開始行 (例: 6)
Private pFirstPageRows As Long      ' 1ページ目の行数 (例: 24 or 41)
Private pNextPageRows As Long       ' 2ページ目以降の行数 (例: 23 or 46)
Private pEndCol As String           ' 印刷終了列 (例: "M")
Private pRefCol As String           ' 最終行判定用の列 (例: "B")

' --- 初期化時にデフォルト値を設定 ---
Private Sub Class_Initialize()
    pStartRow = 6
    pFirstPageRows = 24
    pNextPageRows = 23
    pEndCol = "M"
    pRefCol = "B"
End Sub

' --- プロパティ設定用メソッド ---
Public Sub Configure(firstRows As Long, nextRows As Long, Optional startRow As Long = 6)
    pFirstPageRows = firstRows
    pNextPageRows = nextRows
    pStartRow = startRow
End Sub

'============================================
' メソッド名 : SetPrintArea
' 概要       : 指定シートの印刷範囲を計算して設定する
' 引数       : targetWs As Worksheet
' 戻り値     : Boolean (成功時True)
'============================================
Public Function SetPrintArea(targetWs As Worksheet) As Boolean
    
    Dim lastRow As Long
    Dim dataRows As Long
    Dim pageCount As Long
    Dim printEndRow As Long
    Dim remainingRows As Long
    
    On Error GoTo ErrorHandler
    
    ' 最終行取得
    lastRow = targetWs.Cells(targetWs.Rows.count, pRefCol).End(xlUp).Row
    
    If lastRow < pStartRow Then
        ' データがない場合はエラーとせずFalseを返す（スキップ扱い）
        SetPrintArea = False
        Exit Function
    End If
    
    ' 行数計算
    dataRows = lastRow - (pStartRow - 1)
    
    ' ページ数計算ロジック
    If dataRows <= pFirstPageRows Then
        ' 1ページに収まる場合
        pageCount = 1
        printEndRow = (pStartRow - 1) + pFirstPageRows
    Else
        ' 2ページ以上になる場合
        remainingRows = dataRows - pFirstPageRows
        ' 切り上げ計算: (x + y - 1) \ y
        pageCount = 1 + -Int(-remainingRows / pNextPageRows)
        
        ' 最終印刷行の算出
        printEndRow = (pStartRow - 1) + pFirstPageRows + ((pageCount - 1) * pNextPageRows)
    End If
    
    ' 印刷範囲設定
    targetWs.PageSetup.PrintArea = "B" & pStartRow & ":" & pEndCol & printEndRow
    
    SetPrintArea = True
    Exit Function

ErrorHandler:
    SetPrintArea = False
End Function

'============================================
' メソッド名 : SetFooter
' 概要       : フッターを設定する
' 引数       : targetWs As Worksheet
'            : leftText As String, centerText As String, rightText As String
'            : startPage As Long (開始ページ番号)
'============================================
Public Sub SetFooter(targetWs As Worksheet, leftText As String, centerText As String, rightText As String, startPage As Long)
    With targetWs.PageSetup
        .FirstPageNumber = startPage
        .LeftFooter = leftText
        .CenterFooter = centerText
        .RightFooter = rightText
    End With
End Sub

'============================================
' メソッド名 : GetActualPageCount
' 概要       : 設定された改ページプレビューに基づく実際のページ数を取得
' 引数       : targetWs As Worksheet
'============================================
Public Function GetActualPageCount(targetWs As Worksheet) As Long
    Dim hBreaks As Long, vBreaks As Long
    
    ' 改ページ数を取得（+1でページ数になる）
    hBreaks = targetWs.HPageBreaks.count + 1
    vBreaks = targetWs.VPageBreaks.count + 1
    
    GetActualPageCount = hBreaks * vBreaks
End Function
