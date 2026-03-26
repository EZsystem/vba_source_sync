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

'===========================================================
' クラス名   : acc_clsIcubeTransfer
' 概要       : iCubeデータの履歴統合および関連マスタへの配分
' 依存関係   :
'   - acc_mod_MappingTemplate (テーブル名定数)
'   - com_clsErrorUtility (ログ・エラー管理)
'===========================================================

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
' 処理内容: 重複する枝番工事コードを削除した後、最新データを挿入
'===========================================================
Public Sub ToHistory()
    On Error GoTo Err_Handler
    
    ' 重複削除: at_Icube_累計 から at_Icube に存在するレコードを削除
    m_Db.Execute "DELETE FROM [" & AT_ICUBE_HISTORY & "] " & _
                 "WHERE [枝番工事コード] IN (SELECT [枝番工事コード] FROM [" & AT_ICUBE & "])", dbFailOnError
    
    ' 全件追加: at_Icube から at_Icube_累計 へ全件転写
    m_Db.Execute "INSERT INTO [" & AT_ICUBE_HISTORY & "] SELECT * FROM [" & AT_ICUBE & "]", dbFailOnError
    
    Exit Sub
Err_Handler:
    m_ErrorLog.Notify_Smart_Popup "Class Transfer(History) Error: " & Err.Description
End Sub

'===========================================================
' 関連テーブルへの配分
' 処理内容: ワークおよび累計から各 kt_(現at_) マスタへ差分転写
'===========================================================
Public Sub ToRelatedTables()
    On Error GoTo Err_Handler
    
    ' 転写元テーブルリスト (定数を使用)
    Dim srcTables As Variant
    srcTables = Array(AT_ICUBE, AT_ICUBE_HISTORY)
    
    Dim i As Integer
    For i = 0 To UBound(srcTables)
        ' 各関連テーブルへの転写 (MappingTemplateの定数を使用)
        Call TransferEngine(srcTables(i), AT_KIHON_KANKO, "基本工事コード")
        Call TransferEngine(srcTables(i), AT_KIHON_SAGYO, "基本工事コード")
        Call TransferEngine(srcTables(i), AT_KIHON_JUCHU, "基本工事コード")
        Call TransferEngine(srcTables(i), AT_PROJECT_INFO, "工事コード")
        Call TransferEngine(srcTables(i), AT_EDABAN, "枝番工事コード")
    Next i
    
    Exit Sub
Err_Handler:
    m_ErrorLog.Notify_Smart_Popup "Class Transfer(Related) Error: " & Err.Description
End Sub

'===========================================================
' 内部転写エンジン (Private)
' 処理内容: ターゲットにキーが存在しないレコードのみ、フィールド名が一致する列を転写
'===========================================================
Private Sub TransferEngine(ByVal src As String, ByVal dst As String, ByVal key As String)
    On Error GoTo Err_Handler
    Dim rsS As DAO.Recordset
    Dim rsD As DAO.Recordset
    Dim fld As DAO.Field
    Dim keyVal As String
    
    ' 角括弧で囲み、日本語名やスペースに対応
    Set rsS = m_Db.OpenRecordset(src, dbOpenSnapshot)
    Set rsD = m_Db.OpenRecordset(dst, dbOpenDynaset)

    Do While Not rsS.EOF
        ' キー値のエスケープ処理
        keyVal = Replace(Nz(rsS(key).Value, ""), "'", "''")
        
        ' 存在チェック (未存在の場合のみ AddNew)
        ' DCount内のテーブル名・フィールド名も角括弧で保護
        If DCount("*", "[" & dst & "]", "[" & key & "] = '" & keyVal & "'") = 0 Then
            rsD.AddNew
            For Each fld In rsS.fields
                ' 転送先に同名フィールドが存在するか確認
                If DCount("*", "[" & dst & "]", "[" & key & "] = '" & keyVal & "'") = 0 Then
                    On Error Resume Next
                    ' 個別の値代入エラー（型不一致等）はスキップして継続
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

'===========================================================
' フィールド存在確認ヘルパー
'===========================================================
Private Function FieldExists(ByRef rs As DAO.Recordset, ByVal fName As String) As Boolean
    On Error Resume Next
    Dim f As DAO.Field
    Set f = rs.fields(fName)
    FieldExists = (Err.Number = 0)
    On Error GoTo 0
End Function

