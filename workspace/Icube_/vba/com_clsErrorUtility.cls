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
Private pIsBatchActive As Long       ' 0:通常(ポップアップ有), 1:一括(抑制)
Private pErrorLogs     As Collection ' エラーメッセージ蓄積用

'===========================================================
' サブルーチン名 : Init
' 引数 : isBatch - 一括実行モードにするか (Default: False)
' 機能 : クラスの状態を初期化する。
'===========================================================
Public Sub Init(Optional ByVal isBatch As Boolean = False)
    If isBatch Then
        pIsBatchActive = 1
    Else
        pIsBatchActive = 0
    End If
    Set pErrorLogs = New Collection
End Sub

'===========================================================
' プロパティ Let/Get : IsBatchActive
' 機能 : 実行モードを動的に切り替える。
'===========================================================
Public Property Let IsBatchActive(ByVal val As Boolean)
    If val Then pIsBatchActive = 1 Else pIsBatchActive = 0
End Property

Public Property Get IsBatchActive() As Boolean
    IsBatchActive = (pIsBatchActive = 1)
End Property

'===========================================================
' サブルーチン名 : Notify_Smart_Popup
' 引数 : msg   - 通知内容
'        title - ポップアップのタイトル
'        level - vbInformation, vbExclamation, vbCritical 等
' 機能 : モードに応じて即時ポップアップ、またはログ蓄積を行う。
'===========================================================
Public Sub Notify_Smart_Popup(ByVal msg As String, _
                             Optional ByVal title As String = "System Notification", _
                             Optional ByVal level As VbMsgBoxStyle = vbInformation)
    
    ' ログには常に記録
    pErrorLogs.Add "[" & Format(Now, "hh:nn:ss") & "] " & msg
    
    ' バッチモードでない場合のみポップアップを表示
    If pIsBatchActive = 0 Then
        MsgBox msg, level, title
    Else
        ' バッチモード時はデバッグウィンドウに出力して進行を止めない
        Debug.Print "Batch Log: " & msg
    End If
End Sub

'===========================================================
' 関数名 : Get_Summary
' 戻り値 : 蓄積された全ログの結合文字列
' 機能 : 蓄積されたログを改行区切りで返す。
'===========================================================
Public Function Get_Summary() As String
    If pErrorLogs.Count = 0 Then
        Get_Summary = "No errors or notifications logged."
        Exit Function
    End If
    
    Dim i As Long
    Dim summary As String
    For i = 1 To pErrorLogs.Count
        summary = summary & pErrorLogs(i) & vbCrLf
    Next i
    Get_Summary = summary
End Function

'===========================================================
' サブルーチン名 : Show_Final_Report
' 機能 : 蓄積されたログがあれば最後に一括で表示する。
'===========================================================
Public Sub Show_Final_Report()
    If pErrorLogs.Count > 0 And pIsBatchActive = 1 Then
        MsgBox "バッチ処理中に以下の通知が発生しました：" & vbCrLf & vbCrLf & Get_Summary(), _
               vbInformation, "処理完了レポート"
    End If
End Sub

'===========================================================
' サブルーチン名 : Clear_Logs
' 機能 : 蓄積されたログを消去する。
'===========================================================
Public Sub Clear_Logs()
    Set pErrorLogs = New Collection
End Sub
