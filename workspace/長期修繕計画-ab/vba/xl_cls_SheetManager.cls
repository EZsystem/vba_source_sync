VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_SheetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: cls_SheetManager
' 説明　：シートの生成・削除・管理を行うクラス
' 作成日：2025/11/14
'-------------------------------------
Option Explicit

' --- メンバ変数 ---
Private pIdTable As ListObject
Private pIdColIdx As Long

'============================================
' メソッド名 : Initialize
' 概要       : 処理対象のIDテーブルを設定する
' 引数       : tbl As ListObject - tbl_内訳ID
'============================================
Public Sub Initialize(tbl As ListObject)
    Set pIdTable = tbl
    ' 列インデックス取得
    On Error Resume Next
    pIdColIdx = pIdTable.ListColumns("内訳ID").Index
    On Error GoTo 0
    
    If pIdColIdx = 0 Then
        Err.Raise 1001, "cls_SheetManager", "内訳ID列が見つかりません"
    End If
End Sub

'============================================
' メソッド名 : CreateSheets
' 概要       : テンプレートを複製してシートを量産する
' 引数       : templateName As String - テンプレートシート名
'============================================
Public Sub CreateSheets(templateName As String)
    
    Dim tmplWs As Worksheet
    Dim cell As Range
    Dim newName As String
    Dim createdCnt As Long, skipCnt As Long
    
    ' テンプレート確認
    If Not SheetExists(templateName) Then
        MsgBox "テンプレートシートが見つかりません: " & templateName, vbCritical
        Exit Sub
    End If
    Set tmplWs = ThisWorkbook.Worksheets(templateName)
    
    ' 高速化開始
    ToggleOptimization True
    
    On Error Resume Next
    For Each cell In pIdTable.ListColumns(pIdColIdx).DataBodyRange
        newName = Trim(CStr(cell.value))
        
        If newName <> "" Then
            If SheetExists(newName) Then
                skipCnt = skipCnt + 1
            Else
                ' シートコピー
                tmplWs.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
                With ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
                    .Name = newName
                    ' 必要ならここに初期化処理などを追記可能
                End With
                createdCnt = createdCnt + 1
            End If
        End If
    Next cell
    On Error GoTo 0
    
    ' 高速化終了
    ToggleOptimization False
    
    MsgBox "作成完了: " & createdCnt & "件" & vbCrLf & "スキップ: " & skipCnt & "件", vbInformation
    
End Sub

'============================================
' メソッド名 : DeleteSheets
' 概要       : ID一覧にあるシートを削除する
'============================================
Public Sub DeleteSheets()
    
    Dim cell As Range
    Dim targetName As String
    Dim delCnt As Long, skipCnt As Long
    Dim confirm As VbMsgBoxResult
    
    ' 確認ダイアログ
    confirm = MsgBox("内訳IDテーブルに基づきシートを削除します。" & vbCrLf & "よろしいですか？", _
                     vbYesNo + vbExclamation, "削除確認")
    If confirm = vbNo Then Exit Sub
    
    ToggleOptimization True
    
    On Error Resume Next
    For Each cell In pIdTable.ListColumns(pIdColIdx).DataBodyRange
        targetName = Trim(CStr(cell.value))
        
        If targetName <> "" And SheetExists(targetName) Then
            ' 保護シートチェック（削除禁止）
            If IsProtectedSheet(targetName) Then
                skipCnt = skipCnt + 1
            Else
                ThisWorkbook.Worksheets(targetName).Delete
                delCnt = delCnt + 1
            End If
        Else
            skipCnt = skipCnt + 1
        End If
    Next cell
    On Error GoTo 0
    
    ToggleOptimization False
    
    MsgBox "削除完了: " & delCnt & "件" & vbCrLf & "スキップ: " & skipCnt & "件", vbInformation
    
End Sub

' --- 内部ヘルパー ---

' シート存在チェック
Private Function SheetExists(sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' 削除保護シート判定
Private Function IsProtectedSheet(sName As String) As Boolean
    Select Case sName
        Case "分類", "内訳", "計画表", "出力内訳og", "出力内訳og2"
            IsProtectedSheet = True
        Case Else
            IsProtectedSheet = False
    End Select
End Function

' 高速化設定切替
Private Sub ToggleOptimization(isStart As Boolean)
    With Application
        .ScreenUpdating = Not isStart
        .EnableEvents = Not isStart
        .DisplayAlerts = Not isStart ' 削除時の警告抑制
        If isStart Then
            .Calculation = xlCalculationManual
        Else
            .Calculation = xlCalculationAutomatic
        End If
    End With
End Sub

