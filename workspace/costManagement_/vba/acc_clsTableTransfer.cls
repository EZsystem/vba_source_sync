VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------
' Class: acc_clsTableTransfer
' 説明　：テーブル間のデータ転送と型変換を制御するクラス
'-------------------------------------
Option Compare Database
Option Explicit

Private pSourceTable As String
Private pDestTable As String
Private pCriteria As String
Private pDefaultDict As Object
Private pCleaner As acc_clsDataCleaner


Public Sub Init(ByVal sourceName As String, ByVal destName As String)
    pSourceTable = sourceName
    pDestTable = destName
    Set pCleaner = New acc_clsDataCleaner
    Set pDefaultDict = CreateObject("Scripting.Dictionary")
End Sub

' 抽出条件の設定
Public Property Let Criteria(ByVal value As String)
    pCriteria = value
End Property


Public Property Set DefaultRules(ByRef dict As Object)
    Set pDefaultDict = dict
End Property



Private Sub Class_Terminate()
    Set pCleaner = Nothing
    Set pDefaultDict = Nothing
End Sub

'-------------------------------------
' Class: acc_clsTableTransfer
' 転送実行メソッドの修正版
'-------------------------------------
Public Sub ExecuteTransfer()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsSrc As DAO.Recordset
    Dim rsDst As DAO.Recordset
    Dim fld As DAO.Field
    Dim sql As String
    
    sql = "SELECT * FROM [" & pSourceTable & "]"
    If pCriteria <> "" Then sql = sql & " WHERE " & pCriteria
    
    Set rsSrc = db.OpenRecordset(sql, dbOpenSnapshot)
    Set rsDst = db.OpenRecordset(pDestTable, dbOpenDynaset)
    
    Do Until rsSrc.EOF
        rsDst.AddNew
        For Each fld In rsDst.Fields
            ' --- 【重要】オートナンバー型への書き込みはスキップする ---
            If (fld.Attributes And dbAutoIncrField) = 0 Then
                If FieldExists(rsSrc, fld.Name) Then
                    Dim valRaw As Variant: valRaw = rsSrc(fld.Name).value
                    
                    ' ルール辞書に基づいて値を変換
                    If pDefaultDict.Exists(fld.Name) Then
                        Dim rule As Variant: rule = pDefaultDict(fld.Name)
                        ' rule(0) や rule(1) が Null でも ConvertValue 側で処理
                        rsDst(fld.Name).value = ConvertValue(valRaw, rule(0), rule(1), pCleaner)
                    Else
                        rsDst(fld.Name).value = valRaw
                    End If
                End If
            End If
        Next
        rsDst.Update
        rsSrc.MoveNext
    Loop
    
    rsSrc.Close: rsDst.Close
End Sub
