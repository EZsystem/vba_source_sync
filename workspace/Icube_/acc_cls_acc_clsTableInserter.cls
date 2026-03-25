VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableInserter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsTableInserter
' 説明：転写元から転写先へ、重複（キー一致）を除いた新規レコードのみを挿入する。
'       INSERT INTO ... SELECT ... WHERE NOT IN 構成。
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
' 機能 : 依存オブジェクトと設定リストを初期化する。
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
' 引数 : srcTable - 挿入元テーブル名
'        tgtTable - 挿入先（本番）テーブル名
'        keyField - 重複判定に使用するキー列名
' 機能 : 挿入設定を登録する。
'===========================================================
Public Sub AddSetting(ByVal srcTable As String, ByVal tgtTable As String, ByVal keyField As String)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict("Source") = srcTable
    dict("Target") = tgtTable
    dict("Key") = keyField
    pSettings.Add dict
End Sub

'===========================================================
' サブルーチン名 : InsertAll
' 機能 : 登録されたすべての設定に基づき、新規レコードの挿入を実行する。
'===========================================================
Public Sub InsertAll()
    If pSettings.count = 0 Then Exit Sub
    
    Dim setting As Variant
    For Each setting In pSettings
        InsertNewRecords setting("Source"), setting("Target"), setting("Key")
    Next setting
    
    pErr.Notify_Smart_Popup "すべての新規レコード挿入が完了しました。", "Insert Complete"
End Sub

'===========================================================
' 内部サブルーチン : InsertNewRecords
' 機能 : 共通フィールドを自動抽出し、キーが未存在の行だけを INSERT する。
'===========================================================
Private Sub InsertNewRecords(ByVal src As String, ByVal tgt As String, ByVal key As String)
    On Error GoTo Err_Handler
    
    ' 共通フィールドの取得（Transferクラスと同等のロジック）
    Dim commonFields As String
    commonFields = GetCommonFieldsString(src, tgt)
    
    If commonFields = "" Then
        Err.Raise 1004, , src & " と " & tgt & " の間に共通のフィールドが見つかりません。"
    End If

    ' SQL文の構築：Targetに存在しないKeyのみをSourceから挿入
    Dim sql As String
    sql = "INSERT INTO [" & tgt & "] (" & commonFields & ") " & _
          "SELECT " & commonFields & " FROM [" & src & "] " & _
          "WHERE [" & key & "] NOT IN (SELECT [" & key & "] FROM [" & tgt & "] WHERE [" & key & "] IS NOT NULL) " & _
          "AND [" & key & "] IS NOT NULL;"
          
    pConn.ExecuteSQL sql
    Exit Sub

Err_Handler:
    pErr.Notify_Smart_Popup "挿入失敗 [" & src & " -> " & tgt & "]: " & Err.description, "Insert Error", vbCritical
    Err.Raise Err.Number, "acc_clsTableInserter", Err.description
End Sub

'===========================================================
' 内部関数 : GetCommonFieldsString
' 機能 : 2つのテーブルで共通するフィールド名をカンマ区切りで返す。
'===========================================================
Private Function GetCommonFieldsString(ByVal tbl1 As String, ByVal tbl2 As String) As String
    Dim flds1 As DAO.Fields: Set flds1 = pConn.db.TableDefs(tbl1).Fields
    Dim flds2 As DAO.Fields: Set flds2 = pConn.db.TableDefs(tbl2).Fields
    
    Dim f1 As DAO.Field, f2 As DAO.Field
    Dim result As String
    
    For Each f1 In flds1
        For Each f2 In flds2
            If f1.Name = f2.Name Then
                result = result & "[" & f1.Name & "], "
                Exit For
            End If
        Next f2
    Next f1
    
    If result <> "" Then
        GetCommonFieldsString = Left(result, Len(result) - 2)
    End If
End Function


