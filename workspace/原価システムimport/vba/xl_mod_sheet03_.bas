Attribute VB_Name = "sheet03_"
Option Explicit

Sub sh03_mother01()
'シート：G3のフィルター解除
    Call RemoveFilter_G3_ErrorData
'シート：G3既存データクリア
    Call sh03_clearRange
'原価システムの本ファイルへの取込み忘れチェック
    Call sh03_arr3_mother01
'シート：G2のフィルター解除
    Call RemoveFilter_G2_ErrorData
'エラー一覧のタイトル
    Call sh03arr2_mother01
'エラー一覧
    Call sh03arr_mother01
'シート：G3のフィルター設定
    Call sheet03_filtering
'シート：G2のフィルター設定
    Call sheet02_filtering
End Sub


Sub sh03_mother02()
Call sh03_mother01
Call JumpToCellInSheetG3
End Sub


' ユーザー: 他のシート（G3）のA1セルにジャンプするマクロ
Sub JumpToCellInSheetG3()
    Dim targetSheetName As String
    Dim targetCell As Range
    
    ' 初期設定
    targetSheetName = "G3_原価Sエラー調査"    ' ジャンプするシートの名前
    Set targetCell = Sheets(targetSheetName).Range("A1")   ' ジャンプするセルの設定
    
    ' 対象セルにジャンプ
    Application.GoTo Reference:=targetCell, Scroll:=True
End Sub


' ユーザー: 他のシート（S1）のA1セルにジャンプするマクロ
Sub JumpToCellInSheetS1()
    Dim targetSheetName As String
    Dim targetCell As Range
    
    ' 初期設定
    targetSheetName = "S1_受注、完工、既払い"    ' ジャンプするシートの名前
    Set targetCell = Sheets(targetSheetName).Range("A1")   ' ジャンプするセルの設定
    
    ' 対象セルにジャンプ
    Application.GoTo Reference:=targetCell, Scroll:=True
End Sub
