Attribute VB_Name = "acc_mod_Icube_Transfer"
'Attribute VB_Name = "acc_mod_Icube_Transfer"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : acc_mod_Icube_Transfer
' 概要           : バリデーション済みデータの統合および関連テーブルへの動的転写
' 依存コンポーネント:
'   - acc_mod_MappingTemplate (テーブル名定数)
'   - com_clsErrorUtility (共通エラー/ログ管理)
'===================================================================================================

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Transfer_Icube_To_History
' 概要           : 作業用テーブルから累積テーブルへデータを統合します。
'---------------------------------------------------------------------------------------------------
Public Sub Transfer_Icube_To_History(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim DbObj As DAO.Database: Set DbObj = CurrentDb
    
    ' 1. 重複削除：定数を使用して、枝番工事コードが一致する既存レコードを削除
    Dim SqlDelete As String
    SqlDelete = "DELETE FROM [" & AT_ICUBE_HISTORY & "] WHERE [枝番工事コード] IN (SELECT [枝番工事コード] FROM [" & AT_ICUBE & "])"
    DbObj.Execute SqlDelete, 128 ' 128 = dbFailOnError

    ' 2. 全件追加：最新レコードを履歴に挿入
    Dim SqlInsert As String
    SqlInsert = "INSERT INTO [" & AT_ICUBE_HISTORY & "] SELECT * FROM [" & AT_ICUBE & "]"
    DbObj.Execute SqlInsert, 128
    
    Exit Sub
Err_Handler:
    ErrorLog.Notify_Smart_Popup "Transfer History Error: " & Err.Description, "Transfer Error", vbCritical
End Sub

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Transfer_To_Related_Tables
' 概要           : 各関連マスタテーブルへデータを配分します。
'---------------------------------------------------------------------------------------------------
Public Sub Transfer_To_Related_Tables(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    
    Dim SourceTables As Variant
    ' 定数 AT_ICUBE_HISTORY が正しく定義されていればここでエラーは出ません
    SourceTables = Array(AT_ICUBE, AT_ICUBE_HISTORY)
    
    Dim i As Integer
    Dim SrcName As String
    
    For i = 0 To UBound(SourceTables)
        SrcName = SourceTables(i)
        
        ' 各関連テーブル（定数名を MappingTemplate に合わせる）
        Call TransferTable_Generic(SrcName, AT_KIHON_KANKO, "基本工事コード", ErrorLog)
        Call TransferTable_Generic(SrcName, AT_KIHON_SAGYO, "基本工事コード", ErrorLog)
        Call TransferTable_Generic(SrcName, AT_KIHON_JUCHU, "基本工事コード", ErrorLog)
        
        ' --- 以下の2行を修正 ---
        Call TransferTable_Generic(SrcName, AT_PROJECT_INFO, "工事コード", ErrorLog)
        Call TransferTable_Generic(SrcName, AT_EDABAN, "枝番工事コード", ErrorLog)
    Next i
    
    Exit Sub
Err_Handler:
    ErrorLog.Notify_Smart_Popup "Related Tables Distribution Error: " & Err.Description
End Sub


'---------------------------------------------------------------------------------------------------
' プロシージャ名 : TransferTable_Generic (汎用転写エンジン)
' 概要           : キーの重複をチェックしながら動的に転写。型不一致やオートナンバーに対応。
'---------------------------------------------------------------------------------------------------
Private Sub TransferTable_Generic(ByVal SrcTable As String, ByVal DstTable As String, ByVal KeyField As String, ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim DbObj As DAO.Database: Set DbObj = CurrentDb
    Dim rsSrc As DAO.Recordset
    Dim RsDst As DAO.Recordset
    Dim fld As DAO.Field
    Dim criteria As String
    Dim IsNumericKey As Boolean

    Set rsSrc = DbObj.OpenRecordset("SELECT * FROM [" & SrcTable & "]", 1) ' 1 = dbOpenSnapshot
    Set RsDst = DbObj.OpenRecordset("[" & DstTable & "]", 2) ' 2 = dbOpenDynaset

    ' キーフィールドの型判定 (数値型かテキスト型か)
    IsNumericKey = (rsSrc.Fields(KeyField).Type <> 10) ' 10 = dbText 以外を数値/日付扱いとする簡易判定

    Do While Not rsSrc.EOF
        ' キーが空の場合はスキップ
        If IsNull(rsSrc(KeyField).Value) Then GoTo NextRecord
        
        ' 抽出条件の組み立て
        If IsNumericKey Then
            criteria = "[" & KeyField & "] = " & rsSrc(KeyField).Value
        Else
            criteria = "[" & KeyField & "] = '" & Replace(rsSrc(KeyField).Value, "'", "''") & "'"
        End If
        
        ' 未存在時のみ追加
        If DCount("*", "[" & DstTable & "]", criteria) = 0 Then
            RsDst.AddNew
            For Each fld In rsSrc.Fields
                ' 転送先にフィールドが存在し、かつオートナンバー型(16)でない場合のみ転写
                If Internal_FieldExists(RsDst, fld.Name) Then
                    If (RsDst.Fields(fld.Name).Attributes And 16) = 0 Then
                        On Error Resume Next
                        RsDst(fld.Name).Value = fld.Value
                        If Err.Number <> 0 Then
                            ' 特定フィールドの転送失敗はログ出力して継続
                            Debug.Print "Field Skip: " & DstTable & "." & fld.Name & " (" & Err.Description & ")"
                            Err.Clear
                        End If
                        On Error GoTo Err_Handler
                    End If
                End If
            Next fld
            RsDst.Update
        End If
        
NextRecord:
        rsSrc.MoveNext
    Loop
    
    rsSrc.Close: RsDst.Close
    Exit Sub
Err_Handler:
    ErrorLog.Notify_Smart_Popup DstTable & " Transfer Engine Error: " & Err.Description
End Sub

'---------------------------------------------------------------------------------------------------
' 関数名 : Internal_FieldExists
'---------------------------------------------------------------------------------------------------
Private Function Internal_FieldExists(ByRef RsObj As DAO.Recordset, ByVal FieldName As String) As Boolean
    On Error Resume Next
    Dim FldObj As DAO.Field
    Set FldObj = RsObj.Fields(FieldName)
    Internal_FieldExists = (Err.Number = 0)
    On Error GoTo 0
End Function
