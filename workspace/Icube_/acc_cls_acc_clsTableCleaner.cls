VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsTableCleaner
' 説明：転写元のキーに基づき、転写先から対象行を削除する。
'       主にインポート前の重複排除やデータクリーニングに使用。
' 更新日：2026/03/24（MyLauncher統合・DI対応版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pConn      As acc_clsDbConnector
Private pErr       As com_clsErrorUtility
Private pSettings  As Collection

'===========================================================
' サブルーチン名 : Init
' 引数 : connector - acc_clsDbConnector インスタンス
'        errorUtil - (任意) com_clsErrorUtility
' 機能 : 依存オブジェクトと削除設定リストを初期化する。
'===========================================================
Public Sub Init(ByRef connector As acc_clsDbConnector, Optional ByRef errorUtil As com_clsErrorUtility)
    Set pConn = connector
    
    If errorUtil Is Nothing Then
        Set pErr = New com_clsErrorUtility
        pErr.Init False
    Else
        Set pErr = errorUtil
    End If
    
    Set pSettings = New Collection
End Sub

'===========================================================
' サブルーチン名 : AddSetting
' 引数 : srcTable  - 比較元テーブル名（このテーブルにあるキーが削除対象）
'        tgtTable  - 削除実行先テーブル名
'        keyField  - 比較用キー列名
' 機能 : 削除設定を登録する。
'===========================================================
Public Sub AddSetting(ByVal srcTable As String, ByVal tgtTable As String, ByVal keyField As String)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict("Source") = srcTable
    dict("Target") = tgtTable
    dict("Key") = keyField
    pSettings.Add dict
End Sub

'===========================================================
' サブルーチン名 : CleanAll
' 機能 : 登録されたすべての設定に基づき削除を実行する。
'===========================================================
Public Sub CleanAll()
    If pSettings.count = 0 Then Exit Sub
    
    Dim setting As Variant
    For Each setting In pSettings
        DeleteTargetByKey setting("Source"), setting("Target"), setting("Key")
    Next setting
    
    pErr.Notify_Smart_Popup "すべてのクレンジング（削除）が完了しました。", "Clean Complete"
End Sub

'===========================================================
' 内部サブルーチン : DeleteTargetByKey
' 機能 : SQLの DELETE ... WHERE ... IN 句を用いて高速に削除する。
'===========================================================
Private Sub DeleteTargetByKey(ByVal src As String, ByVal tgt As String, ByVal key As String)
    On Error GoTo Err_Handler
    
    Dim sql As String
    ' キーが Null でないレコードを対象に、一致するものを削除
    sql = "DELETE FROM [" & tgt & "] " & _
          "WHERE [" & key & "] IN (" & _
          "SELECT [" & key & "] FROM [" & src & "] WHERE [" & key & "] IS NOT NULL);"
          
    pConn.ExecuteSQL sql
    Exit Sub

Err_Handler:
    pErr.Notify_Smart_Popup "削除失敗 [" & tgt & "]: " & Err.description, "Delete Error", vbCritical
    ' トランザクション中断のため再スロー
    Err.Raise Err.Number, "acc_clsTableCleaner", Err.description
End Sub

'===========================================================
' サブルーチン名 : ClearAllData
' 引数 : tableName - 全件削除するテーブル名
' 機能 : 指定テーブルの全レコードを削除する。
'===========================================================
Public Sub ClearAllData(ByVal TableName As String)
    On Error GoTo Err_Handler
    pConn.ExecuteSQL "DELETE FROM [" & TableName & "];"
    Exit Sub

Err_Handler:
    pErr.Notify_Smart_Popup "全件削除失敗 [" & TableName & "]: " & Err.description, "Clear Error", vbCritical
    Err.Raise Err.Number
End Sub
