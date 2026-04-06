Attribute VB_Name = "xl_modGenkaBatchRunner"
'-------------------------------------
' Module: xl_modGenkaBatchRunner
' 説明 : 原価系のAccess処理をExcelから一括実行する
' 作成日: 2025/08/18
'-------------------------------------

Option Explicit

'============================================
' プロシージャ名 : Genka_RunAll
' 概要          : 1→5 をまとめて実行。途中エラーで停止して通知する
'============================================
Public Sub Genka_RunAll()

    Dim accApp As Object          ' Access.Application
    Dim dbPath As String
    Dim stepName As String
    Dim stopped As Boolean
    
    On Error GoTo EH

    dbPath = GetAccdbPath_()
    If Len(dbPath) = 0 Then Err.Raise vbObjectError + 700, , "Accessパスが未設定"
    If Dir(dbPath) = "" Then Err.Raise vbObjectError + 701, , "Accessファイルが見つからない: " & dbPath

    Set accApp = CreateObject("Access.Application")
    accApp.OpenCurrentDatabase dbPath
    
    ' --- 1. 工事原価�@：インポート（枝番→基本） ---
    stepName = "1: Import (Edaban→Kihon)"
    LogRow_ stepName, "START", ""
    accApp.Run "Transfer_genkaEdaban_Import"
    accApp.Run "Transfer_genkaKihon_Import"
    LogRow_ stepName, "OK", ""
    
    ' --- 2. 工事原価�@：枝番コード更新＋不一致チェック ---
    stepName = "2: Update/Check EdabanCode"
    LogRow_ stepName, "START", ""
    accApp.Run "UpdateEdabanCodeByReference"
    accApp.Run "CheckEdabanCodeError"
    LogRow_ stepName, "OK", ""
    
    ' --- 3. 工事原価�A：基本エラーチェック（選択クエリ） ---
    stepName = "3: Kihon Error Check (Query)"
    LogRow_ stepName, "START", ""
    accApp.DoCmd.OpenQuery "sel_原価S基本工事errCheck", 2
    accApp.DoCmd.Close 5, "sel_原価S基本工事errCheck"
    LogRow_ stepName, "OK", ""
    
    ' --- 4. 工事原価�A：枝番エラーチェック�@（選択クエリ） ---
    stepName = "4: Edaban Error Check 1 (Query)"
    LogRow_ stepName, "START", ""
    accApp.DoCmd.OpenQuery "sel_原価S枝番工事errCheck1", 2
    accApp.DoCmd.Close 5, "sel_原価S枝番工事errCheck1"
    LogRow_ stepName, "OK", ""
    
    ' --- 5. 工事原価�A：枝番エラーチェック�A（選択クエリ） ---
    stepName = "5: Edaban Error Check 2 (Query)"
    LogRow_ stepName, "START", ""
    accApp.DoCmd.OpenQuery "sel_原価S枝番工事errCheck2", 2
    accApp.DoCmd.Close 5, "sel_原価S枝番工事errCheck2"
    LogRow_ stepName, "OK", ""

    accApp.Quit
    Set accApp = Nothing
    
    MsgBox "一括実行がすべて完了したにゃ", vbInformation, "そうじろう"
    Exit Sub

EH:
    On Error Resume Next
    LogRow_ stepName, "ERR", Err.Description
    If Not accApp Is Nothing Then accApp.Quit
    Set accApp = Nothing
    
    MsgBox "一括実行は途中で停止したにゃ" & vbCrLf & _
           "手順 : " & stepName & vbCrLf & _
           "理由 : " & Err.Description & vbCrLf & vbCrLf & _
           "→ 以降は Access 本体で個別に実施してほしいにゃ", _
           vbExclamation, "そうじろう"
End Sub   ' ← Subの終わり

'============================================
' プロシージャ名 : GetAccdbPath_
' 概要          : 「原価S_temp」C3 から accdb フルパスを取得する
'============================================
Private Function GetAccdbPath_() As String
    Dim ws As Worksheet
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets("原価S_temp")
    GetAccdbPath_ = Trim$(CStr(ws.Range("C3").value))
    Exit Function
EH:
    GetAccdbPath_ = ""
End Function   ' ← Functionの終わり

'============================================
' プロシージャ名 : LogRow_
' 概要          : ログシート「原価S_log」に 〔日時, 手順, 状態, 詳細〕を追記する
'============================================
Private Sub LogRow_(ByVal stepName As String, ByVal status As String, ByVal detail As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    
    Set ws = EnsureLogSheet_()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    If nextRow < 2 Then nextRow = 2
    
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = stepName
    ws.Cells(nextRow, 3).value = status
    ws.Cells(nextRow, 4).value = detail
End Sub   ' ← Subの終わり

'============================================
' プロシージャ名 : EnsureLogSheet_
' 概要          : ログシートを作成／取得する（見出し付き）
'============================================
Private Function EnsureLogSheet_() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("原価S_log")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "原価S_log"
        ws.Range("A1:D1").value = Array("日時", "手順", "状態", "詳細")
        ws.Columns("A:D").EntireColumn.AutoFit
    End If
    
    Set EnsureLogSheet_ = ws
End Function   ' ← Functionの終わり


