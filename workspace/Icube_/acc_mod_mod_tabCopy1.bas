Attribute VB_Name = "mod_tabCopy1"
'-------------------------------------
' Module: acc_mod_IcubeDataTransfer
' 説明  : 仮テーブルから本テーブルへ、条件付き転写を行う処理
' 作成日: 2025/06/02
' 更新日: -
'-------------------------------------
Option Compare Database
Option Explicit

'============================================
' プロシージャ名 : Transfer_IcubeData
' 概要           : 仮テーブル（Icube_）→本テーブル（Icube_累計）へデータ転写
' 引数           : なし
' 呼び出し元     : フォームやマクロ等
' 処理内容       : 本テーブルで「枝番工事コード」が重複するレコードを削除後、仮テーブルの全データを追加する
'============================================
Public Sub Transfer_IcubeData()
    Dim sqlDelete As String
    Dim sqlInsert As String

    ' --- 削除クエリ（同じ枝番工事コードを持つ本テーブルのレコードを削除） ---
    sqlDelete = ""
    sqlDelete = sqlDelete & "DELETE FROM Icube_累計 " & vbCrLf
    sqlDelete = sqlDelete & "WHERE [枝番工事コード] IN (" & vbCrLf
    sqlDelete = sqlDelete & "    SELECT [枝番工事コード] FROM Icube_" & vbCrLf
    sqlDelete = sqlDelete & ")"

    ' --- 挿入クエリ（仮テーブルの全データを追加） ---
    sqlInsert = ""
    sqlInsert = sqlInsert & "INSERT INTO Icube_累計 " & vbCrLf
    sqlInsert = sqlInsert & "SELECT * FROM Icube_"

    ' --- クエリ実行 ---
    On Error GoTo ErrHandler

    CurrentDb.Execute sqlDelete, dbFailOnError
    CurrentDb.Execute sqlInsert, dbFailOnError

    MsgBox "データ転写が完了しましたニャ！", vbInformation

    Exit Sub

ErrHandler:
    MsgBox "エラーが発生したニャ：" & Err.description, vbCritical
End Sub

