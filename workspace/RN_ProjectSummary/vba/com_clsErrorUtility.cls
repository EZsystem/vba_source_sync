VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "com_clsErrorUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: com_clsErrorUtility
' 説明：通知管理およびエラーログ蓄積を行うユーティリティ。
'       一括実行フラグ(pIsBatchActive)により、ポップアップを制御する。
' 更新日：2026/03/24（MyLauncher統合・通知制御対応版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pIsBatchActive As Boolean    ' True:一括(抑制), False:通常(ポップアップ有)
Private pErrorLogs     As Collection ' エラーメッセージ蓄積用

'===========================================================
' 初期化メソッド (共通)
'===========================================================
Public Sub Init(Optional ByVal isBatch As Boolean = False)
    pIsBatchActive = isBatch
    Set pErrorLogs = New Collection
End Sub

' Kenmu_Transcribe 等で期待される別名
Public Sub Initialize(Optional ByVal debugMode As Boolean = False)
    Call Init(isBatch:=debugMode)
End Sub

'===========================================================
' プロパティ
'===========================================================
Public Property Let IsBatchActive(ByVal val As Boolean)
    pIsBatchActive = val
End Property
Public Property Get IsBatchActive() As Boolean
    IsBatchActive = pIsBatchActive
End Property

Public Property Get ErrorCount() As Long
    If pErrorLogs Is Nothing Then ErrorCount = 0 Else ErrorCount = pErrorLogs.count
End Property

'===========================================================
' 通知・ログ記録メソッド
'===========================================================

' 旧来のログ記録メソッド (互換性維持用)
Public Sub LogError(ByVal title As String, ByVal msg As String)
    If pErrorLogs Is Nothing Then Set pErrorLogs = New Collection
    pErrorLogs.Add "[" & Format(Now, "hh:nn:ss") & "] " & title & ": " & msg
    Debug.Print "LogError: " & title & " - " & msg
End Sub

' 新しいスマートポップアップ通知
Public Sub Notify_Smart_Popup(ByVal msg As String, _
                             Optional ByVal title As String = "System Notification", _
                             Optional ByVal level As VbMsgBoxStyle = vbInformation)
    
    If pErrorLogs Is Nothing Then Set pErrorLogs = New Collection
    pErrorLogs.Add "[" & Format(Now, "hh:nn:ss") & "] " & msg
    
    ' バッチモードでない場合のみポップアップを表示
    If Not pIsBatchActive Then
        MsgBox msg, level, title
    Else
        Debug.Print "Batch Log: " & msg
    End If
End Sub

'===========================================================
' ユーティリティ
'===========================================================
Public Function Get_Summary() As String
    If pErrorLogs Is Nothing Or pErrorLogs.count = 0 Then
        Get_Summary = "No errors or notifications logged."
        Exit Function
    End If
    
    Dim i As Long: Dim summary As String
    For i = 1 To pErrorLogs.count
        summary = summary & pErrorLogs(i) & vbCrLf
    Next i
    Get_Summary = summary
End Function

' 旧来の全エラー表示メソッド (互換性維持用)
Public Sub ShowAllErrors(Optional ByVal showInMsgBox As Boolean = True)
    If ErrorCount > 0 Then
        If showInMsgBox Then
            MsgBox "以下のエラー・通知が発生しました：" & vbCrLf & vbCrLf & Get_Summary(), _
                   vbInformation, "エラーログ一括表示"
        Else
            Debug.Print Get_Summary()
        End If
    End If
End Sub

' 最終レポート表示
Public Sub Show_Final_Report()
    If ErrorCount > 0 And pIsBatchActive Then
        MsgBox "バッチ処理中に以下の通知が発生しました：" & vbCrLf & vbCrLf & Get_Summary(), _
               vbInformation, "処理完了レポート"
    End If
End Sub

Public Sub Clear_Logs()
    Set pErrorLogs = New Collection
End Sub
