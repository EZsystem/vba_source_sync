VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xl_clsWorkbookManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: xl_clsWorkbookManager
' 説明：Excelアプリケーションおよびワークブックのライフサイクルを管理する。
'       ブックが1つも開いていない状況での Calculation 操作エラーを回避する。
' 更新日：2026/03/24（MyLauncher統合・堅牢化版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pxlApp          As Object  ' Excel.Application (Late Binding)
Private pWb             As Object  ' Excel.Workbook
Private pErr            As com_clsErrorUtility
Private pIsAppCreated   As Boolean ' 自分がExcelを起動したかどうかのフラグ
Private pOriginalState  As Object  ' 高速化前の状態保存用（Dictionary）

'===========================================================
' サブルーチン名 : Init
' 引数 : errorUtil - com_clsErrorUtility インスタンス
'        showExcel - Excelを表示するか (Default: False)
' 機能 : Excelアプリケーションへの接続を初期化する。
'===========================================================
Public Sub Init(ByRef errorUtil As com_clsErrorUtility, Optional ByVal showExcel As Boolean = False)
    Set pErr = errorUtil
    Set pOriginalState = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    ' 既に起動しているExcelへのアタッチを試行
    Set pxlApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    
    If pxlApp Is Nothing Then
        ' 起動していなければ新規作成
        Set pxlApp = CreateObject("Excel.Application")
        pIsAppCreated = True
    Else
        pIsAppCreated = False
    End If
    
    pxlApp.Visible = showExcel
End Sub

'===========================================================
' プロパティ Let : FastMode
' 引数 : active - Trueで高速化開始, Falseで解除
' 機能 : 画面更新停止、計算手動化、イベント停止を一括制御する。
'        ブックが存在しない場合は Calculation の操作を安全にスキップする。
'===========================================================
Public Property Let FastMode(ByVal active As Boolean)
    If pxlApp Is Nothing Then Exit Property
    
    ' Calculationはブックが開いていないと操作できないため、ブック数を確認
    Dim hasWorkbook As Boolean
    hasWorkbook = (pxlApp.Workbooks.Count > 0)
    
    On Error Resume Next ' 予期せぬCOMエラーに対する最終ガード
    If active Then
        ' 現在の状態を保存（DictionaryのExistsチェックを容易にするため、計算はブックがある場合のみ保存）
        pOriginalState("ScreenUpdating") = pxlApp.ScreenUpdating
        pOriginalState("EnableEvents") = pxlApp.EnableEvents
        If hasWorkbook Then
            pOriginalState("Calculation") = pxlApp.Calculation
        End If
        
        ' 高速化設定
        pxlApp.ScreenUpdating = False
        pxlApp.EnableEvents = False
        If hasWorkbook Then
            pxlApp.Calculation = -4135 ' xlCalculationManual
        End If
    Else
        ' 保存していた状態に復元（Existsメソッドで保存されている場合のみ復元）
        If pOriginalState.Exists("ScreenUpdating") Then pxlApp.ScreenUpdating = pOriginalState("ScreenUpdating")
        If pOriginalState.Exists("EnableEvents") Then pxlApp.EnableEvents = pOriginalState("EnableEvents")
        
        ' 復元時もブックがある場合のみ操作
        If hasWorkbook And pOriginalState.Exists("Calculation") Then
            pxlApp.Calculation = pOriginalState("Calculation")
        End If
        
        ' 復元後は履歴をクリア
        pOriginalState.RemoveAll
    End If
    On Error GoTo 0
End Property

'===========================================================
' 関数名 : OpenWorkbook
' 引数 : path - Excelファイルパス
' 戻り値 : 開いた Workbook オブジェクト
'===========================================================
Public Function OpenWorkbook(ByVal path As String) As Object
    On Error GoTo Err_Handler
    
    If pxlApp Is Nothing Then Err.Raise 1001, , "Excelが初期化されていません。"
    
    Set pWb = pxlApp.Workbooks.Open(path)
    Set OpenWorkbook = pWb
    Exit Function

Err_Handler:
    pErr.Notify_Smart_Popup "ブックを開けませんでした: " & path, "File Open Error", vbCritical
    Set OpenWorkbook = Nothing
End Function

'===========================================================
' サブルーチン名 : CloseWorkbook
' 引数 : saveChanges - 保存して閉じるか
'===========================================================
Public Sub CloseWorkbook(Optional ByVal saveChanges As Boolean = False)
    If Not pWb Is Nothing Then
        pWb.Close saveChanges
        Set pWb = Nothing
    End If
End Sub

'===========================================================
' クリーンアップ : Class_Terminate
' 機能 : 自分が起動したExcelプロセスのみ、安全に終了させる。
'===========================================================
Private Sub Class_Terminate()
    On Error Resume Next
    ' 高速化設定を念のため解除
    Me.FastMode = False
    
    If pIsAppCreated Then
        If Not pxlApp Is Nothing Then
            pxlApp.Quit
        End If
    End If
    Set pxlApp = Nothing
    Set pWb = Nothing
    Set pOriginalState = Nothing
End Sub

