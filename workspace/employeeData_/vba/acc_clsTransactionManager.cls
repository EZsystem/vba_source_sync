VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTransactionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsTransactionManager
' 説明　：DAOトランザクションの開始・コミット・ロールバックを管理するクラス
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
' DAOトランザクションは Workspace 単位で管理するため、Databaseオブジェクトは不要です
Private pInTransaction As Boolean

'=================================================
' サブルーチン名 : Init
'=================================================
Public Sub Init(Optional db As DAO.Database = Nothing)
    pInTransaction = False
End Sub

'=================================================
' サブルーチン名 : BeginTransaction
' 機能 : トランザクションを開始する。DAOの正しい階層（Workspace）を呼び出します
'=================================================
Public Sub BeginTransaction()
    If Not pInTransaction Then
        ' DAOの仕様：トランザクションは Workspace(0) で開始します
        DBEngine.Workspaces(0).BeginTrans
        pInTransaction = True
    End If
End Sub

'=================================================
' サブルーチン名 : Commit
' 機能 : トランザクションをコミットして確定する
'=================================================
Public Sub Commit()
    If pInTransaction Then
        DBEngine.Workspaces(0).CommitTrans
        pInTransaction = False
    End If
End Sub

'=================================================
' サブルーチン名 : Rollback
' 機能 : トランザクションをロールバックして元に戻す
'=================================================
Public Sub Rollback()
    If pInTransaction Then
        ' DAOのメソッド名は Rollback です（RollbackTransではありません）
        DBEngine.Workspaces(0).Rollback
        pInTransaction = False
    End If
End Sub
