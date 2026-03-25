VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTransactionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsTransactionManager
' 説明：DAOトランザクション（Begin/Commit/Rollback）を安全に管理する。
'       二重開始防止および、終了時の自動ロールバック保護機能を内蔵。
' 更新日：2026/03/24（MyLauncher統合・堅牢化版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pWS            As DAO.Workspace
Private pInTransaction As Boolean

'===========================================================
' サブルーチン名 : Init
' 機能 : デフォルトのワークスペースを初期化する。
'===========================================================
Public Sub Init()
    Set pWS = DBEngine.Workspaces(0)
    pInTransaction = False
End Sub

'===========================================================
' プロパティ Get : InTransaction
' 機能 : 現在トランザクション中かどうかを返す。
'===========================================================
Public Property Get InTransaction() As Boolean
    InTransaction = pInTransaction
End Property

'===========================================================
' サブルーチン名 : BeginTransaction
' 機能 : トランザクションを開始する。二重開始は無視する。
'===========================================================
Public Sub BeginTransaction()
    If pWS Is Nothing Then Me.Init
    
    If Not pInTransaction Then
        pWS.BeginTrans
        pInTransaction = True
        Debug.Print "[TX] Transaction Started."
    End If
End Sub

'===========================================================
' サブルーチン名 : Commit
' 機能 : 処理を確定させる。
'===========================================================
Public Sub Commit()
    If pInTransaction Then
        pWS.CommitTrans
        pInTransaction = False
        Debug.Print "[TX] Transaction Committed."
    End If
End Sub

'===========================================================
' サブルーチン名 : Rollback
' 機能 : 処理を破棄し、開始前の状態に戻す。
'===========================================================
Public Sub Rollback()
    If pInTransaction Then
        pWS.Rollback
        pInTransaction = False
        Debug.Print "[TX] Transaction Rolled Back."
    End If
End Sub

'===========================================================
' クリーンアップ : Class_Terminate
' 機能 : コミットされずに終了（異常終了等）した場合、自動でロールバックする。
'===========================================================
Private Sub Class_Terminate()
    ' トランザクションが残ったまま破棄されたらロールバック（安全策）
    If pInTransaction Then
        On Error Resume Next
        pWS.Rollback
        Debug.Print "[TX] Safety Rollback executed in Terminate."
    End If
    Set pWS = Nothing
End Sub
