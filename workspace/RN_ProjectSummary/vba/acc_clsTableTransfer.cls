VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===========================================================
' Class Module: acc_clsTableTransfer
' 説明：テーブル間のデータ転送を一括管理する。
'       1. 高速なSQL一括挿入 (ExecuteTransfer)
'       2. 辞書に基づいた精密な型変換挿入 (ExecuteTransferWithRules)
' 更新日：2026/03/26（RN/原価管理統合・汎用版）
'===========================================================
Option Explicit

Private pDb As DAO.Database
Private pCleaner As acc_clsDataCleaner

'===========================================================
' メソッド : Init
'===========================================================
Public Sub Init(ByRef db As DAO.Database)
    Set pDb = db
    Set pCleaner = New acc_clsDataCleaner
    pCleaner.Init
End Sub

'===========================================================
' メソッド : ExecuteTransferWithRules (★今回追加されたメソッド)
' 概要 : マッピング辞書に基づき、1レコードずつ型変換を行いながら転送する。
' 引数 : src - 元テーブル名, tgt - 先テーブル名, ruleDict - 辞書オブジェクト, criteria - 抽出条件SQL
'===========================================================
Public Sub ExecuteTransferWithRules(ByVal src As String, ByVal tgt As String, ByVal ruleDict As Object, Optional ByVal criteria As String = "")
    Dim rsSrc As DAO.Recordset
    Dim rsTgt As DAO.Recordset
    Dim sql As String
    Dim fldName As Variant
    
    ' 1. ソースのオープン
    sql = "SELECT * FROM [" & src & "]"
    If criteria <> "" Then sql = sql & " WHERE " & criteria
    
    Set rsSrc = pDb.OpenRecordset(sql, dbOpenSnapshot)
    Set rsTgt = pDb.OpenRecordset(tgt, dbOpenDynaset)
    
    ' 2. レコードループ
    Do Until rsSrc.EOF
        rsTgt.AddNew
        For Each fldName In ruleDict.Keys
            ' 転送先と転送元にフィールドが存在するかチェック
            If FieldExistsInRS(rsTgt, CStr(fldName)) Then
                If FieldExistsInRS(rsSrc, CStr(fldName)) Then
                    Dim rawVal As Variant: rawVal = rsSrc(CStr(fldName)).Value
                    Dim rule As Variant: rule = ruleDict(fldName)
                    
                    ' rule(0) はデータ型名、rule(1) はデフォルト値
                    rsTgt(CStr(fldName)).Value = ConvertByRule(rawVal, CStr(rule(0)), rule(1))
                End If
            End If
        Next
        rsTgt.Update
        rsSrc.MoveNext
    Loop
    
    rsSrc.Close
    rsTgt.Close
End Sub

'===========================================================
' メソッド : ExecuteTransfer (RN標準・高速SQL版)
' 概要 : 列名が一致する項目をSQLのINSERT文で一括転送する。
'===========================================================
Public Sub ExecuteTransfer(ByVal SrcTable As String, ByVal tgtTable As String, Optional ByVal KeyField As String = "")
    Dim insertFields As String
    Dim selectFields As String
    
    ' フィールドリストの生成
    GetFieldLists_Internal SrcTable, tgtTable, insertFields, selectFields
    
    If insertFields = "" Then Err.Raise 1004, , "共通列が見つかりません。"

    ' 既存レコードの削除
    Dim sqlDel As String
    If KeyField <> "" Then
        sqlDel = "DELETE FROM [" & tgtTable & "] WHERE [" & KeyField & "] IN (SELECT Val([" & KeyField & "]) FROM [" & SrcTable & "]);"
    Else
        sqlDel = "DELETE FROM [" & tgtTable & "];"
    End If
    pDb.Execute sqlDel, 128 ' dbFailOnError

    ' 一括挿入
    Dim sqlIns As String
    sqlIns = "INSERT INTO [" & tgtTable & "] (" & insertFields & ") SELECT " & selectFields & " FROM [" & SrcTable & "];"
    pDb.Execute sqlIns, 128
End Sub

' --- 内部補助関数：型変換ロジック ---
Private Function ConvertByRule(ByVal val As Variant, ByVal typeName As String, ByVal defVal As Variant) As Variant
    Select Case typeName
        Case "テキスト型":   ConvertByRule = pCleaner.CleanText(val, CStr(Nz(defVal, "")))
        Case "長整数型":     ConvertByRule = pCleaner.CleanLong(val, CLng(Nz(defVal, 0)))
        Case "倍精度型", "倍精度浮動小数点型": ConvertByRule = pCleaner.CleanDouble(val, CDbl(Nz(defVal, 0)))
        Case "通貨型":       ConvertByRule = pCleaner.CleanCurrency(val, CCur(Nz(defVal, 0)))
        Case "日付型", "日付/時刻型": ConvertByRule = pCleaner.CleanDate(val, CDate(Nz(defVal, #12/31/2099#)))
        Case Else:           ConvertByRule = val
    End Select
End Function

' --- 内部補助関数：Recordset内のフィールド存在確認 ---
Private Function FieldExistsInRS(ByRef rs As DAO.Recordset, ByVal fldName As String) As Boolean
    On Error Resume Next
    Dim dummy As Variant: dummy = rs.fields(fldName).Name
    FieldExistsInRS = (Err.Number = 0)
    On Error GoTo 0
End Function

' --- 内部補助関数：SQL用フィールドリスト作成 ---
Private Sub GetFieldLists_Internal(ByVal src As String, ByVal tgt As String, ByRef outInsert As String, ByRef outSelect As String)
    Dim fldsSrc As DAO.fields: Set fldsSrc = pDb.TableDefs(src).fields
    Dim fldsTgt As DAO.fields: Set fldsTgt = pDb.TableDefs(tgt).fields
    Dim fSrc As DAO.Field, fTgt As DAO.Field
    Dim insList As String: insList = ""
    Dim selList As String: selList = ""
    
    For Each fSrc In fldsSrc
        If (fSrc.Attributes And 16) = 0 Then ' オートナンバー除外
            For Each fTgt In fldsTgt
                If StrComp(fSrc.Name, fTgt.Name, 1) = 0 Then
                    insList = insList & "[" & fTgt.Name & "], "
                    selList = selList & "IIf([" & fSrc.Name & "]='', Null, [" & fSrc.Name & "]), "
                    Exit For
                End If
            Next fTgt
        End If
    Next fSrc
    
    If insList <> "" Then
        outInsert = Left(insList, Len(insList) - 2)
        outSelect = Left(selList, Len(selList) - 2)
    End If
End Sub
