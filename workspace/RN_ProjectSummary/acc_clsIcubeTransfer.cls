VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsIcubeTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private m_ErrorLog As com_clsErrorUtility
Private m_Db As DAO.Database

'===========================================================
' 初期化メソッド
'===========================================================
Public Sub Init(ByRef ErrorLog As com_clsErrorUtility)
    Set m_ErrorLog = ErrorLog
    Set m_Db = CurrentDb
End Sub

'===========================================================
' 累計テーブルへの統合
'===========================================================
Public Sub ToHistory()
    On Error GoTo Err_Handler
    
    ' 重複削除
    m_Db.Execute "DELETE FROM Icube_累計 WHERE [枝番工事コード] IN (SELECT [枝番工事コード] FROM Icube_)", dbFailOnError
    ' 全件追加
    m_Db.Execute "INSERT INTO Icube_累計 SELECT * FROM Icube_", dbFailOnError
    
    Exit Sub
Err_Handler:
    m_ErrorLog.Notify_Smart_Popup "Class Transfer(History) Error: " & Err.Description
End Sub

'===========================================================
' 関連テーブル(kt_群)への配分
'===========================================================
Public Sub ToRelatedTables()
    On Error GoTo Err_Handler
    
    Dim srcTables As Variant
    Dim i As Integer
    srcTables = Array("Icube_", "Icube_累計")
    
    For i = 0 To UBound(srcTables)
        ' 各テーブルへの転写を汎用エンジンで実行
        Call TransferEngine(srcTables(i), "kt_基本工事_完工", "基本工事コード")
        Call TransferEngine(srcTables(i), "kt_基本工事_作業所", "基本工事コード")
        Call TransferEngine(srcTables(i), "kt_基本工事_受注", "基本工事コード")
        Call TransferEngine(srcTables(i), "kt_工事コード情報", "工事コード")
        Call TransferEngine(srcTables(i), "kt_枝番工事", "枝番工事コード")
    Next i
    
    Exit Sub
Err_Handler:
    m_ErrorLog.Notify_Smart_Popup "Class Transfer(Related) Error: " & Err.Description
End Sub

'===========================================================
' 内部転写エンジン (Private)
'===========================================================
Private Sub TransferEngine(ByVal src As String, ByVal dst As String, ByVal key As String)
    On Error GoTo Err_Handler
    Dim rsS As DAO.Recordset
    Dim rsD As DAO.Recordset
    Dim fld As DAO.Field
    Dim keyVal As String
    
    Set rsS = m_Db.OpenRecordset("SELECT * FROM [" & src & "]", dbOpenSnapshot)
    Set rsD = m_Db.OpenRecordset(dst, dbOpenDynaset)

    Do While Not rsS.EOF
        keyVal = Replace(Nz(rsS(key).Value, ""), "'", "''")
        
        ' 存在チェック (未存在の場合のみ)
        If DCount("*", "[" & dst & "]", "[" & key & "] = '" & keyVal & "'") = 0 Then
            rsD.AddNew
            For Each fld In rsS.fields
                If FieldExists(rsD, fld.Name) Then
                    On Error Resume Next
                    rsD(fld.Name).Value = fld.Value
                    Err.Clear
                    On Error GoTo Err_Handler
                End If
            Next fld
            rsD.Update
        End If
        rsS.MoveNext
    Loop
    
    rsS.Close
    rsD.Close
    Exit Sub
Err_Handler:
    m_ErrorLog.Notify_Smart_Popup "Engine Error (" & dst & "): " & Err.Description
End Sub

' フィールド存在確認ヘルパー
Private Function FieldExists(ByRef rs As DAO.Recordset, ByVal fName As String) As Boolean
    On Error Resume Next
    Dim f As DAO.Field
    Set f = rs.fields(fName)
    FieldExists = (Err.Number = 0)
    On Error GoTo 0
End Function
