VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsTableUpdater (EZsystem Integrated)
' 説明　：テーブルのフィールド演算およびマスタ同期(Upsert)を提供するクラス
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private pDB            As DAO.Database

' --- 公開プロパティ (マスタ同期用) ---
Public TargetTable     As String    ' 本番テーブル名
Public TempTable       As String    ' 一時テーブル名
Public KeyField        As String    ' 主キー（社員番号など）
Public FieldMapping    As Variant   ' 更新対象フィールドの配列

'=================================================
' サブルーチン名 : Init
'=================================================
Public Sub Init(Optional tableName As String = "")
    Set pDB = CurrentDb
    If tableName <> "" Then TargetTable = tableName
End Sub

'=================================================
' 関数名 : ImportFromBuffer
' 機能 : 文字列バッファを分解してTempテーブルへロードする
'=================================================
Public Function ImportFromBuffer(rawData As String, rowDelim As String, colDelim As String) As Boolean
    On Error GoTo ErrLine
    Dim lines As Variant, cols As Variant
    Dim rs As DAO.Recordset
    Dim i As Long, j As Long
    
    If TempTable = "" Then Err.Raise 999, , "TempTableが指定されていません"
    
    ' 一時テーブルのクリア
    pDB.Execute "DELETE * FROM [" & TempTable & "];", dbFailOnError
    
    Set rs = pDB.OpenRecordset(TempTable, dbOpenDynaset)
    lines = Split(rawData, rowDelim)
    
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then
            cols = Split(lines(i), colDelim)
            rs.AddNew
            ' 社員番号(Key) + マッピング列の順で流し込み
            rs.Fields(KeyField).value = Trim(cols(0))
            For j = 0 To UBound(FieldMapping)
                If (j + 1) <= UBound(cols) Then
                    rs.Fields(FieldMapping(j)).value = Trim(cols(j + 1))
                End If
            Next j
            rs.Update
        End If
    Next i
    rs.Close
    ImportFromBuffer = True
    Exit Function
ErrLine:
    ImportFromBuffer = False
End Function

'=================================================
' サブルーチン名 : SyncWithMaster
' 機能 : Tempと本番を比較し、更新(UPDATE)と追加(INSERT)を一括実行する
'=================================================
Public Sub SyncWithMaster()
    Dim sqlUpdate As String, sqlInsert As String
    Dim fld As Variant, setClause As String, fldList As String, valList As String

    ' 1. UPDATE文の構築 (既存レコードの更新)
    For Each fld In FieldMapping
        setClause = setClause & "[" & TargetTable & "].[" & fld & "] = [" & TempTable & "].[" & fld & "], "
    Next fld
    setClause = Left(setClause, Len(setClause) - 2)
    
    sqlUpdate = "UPDATE [" & TargetTable & "] INNER JOIN [" & TempTable & "] ON [" & TargetTable & "].[" & KeyField & "] = [" & TempTable & "].[" & KeyField & "] " & _
                "SET " & setClause & ";"
    pDB.Execute sqlUpdate, dbFailOnError

    ' 2. INSERT文の構築 (新規レコードの追加)
    fldList = "[" & KeyField & "], "
    valList = "[" & TempTable & "].[" & KeyField & "], "
    For Each fld In FieldMapping
        fldList = fldList & "[" & fld & "], "
        valList = valList & "[" & TempTable & "].[" & fld & "], "
    Next fld
    fldList = Left(fldList, Len(fldList) - 2)
    valList = Left(valList, Len(valList) - 2)

    sqlInsert = "INSERT INTO [" & TargetTable & "] (" & fldList & ") " & _
                "SELECT " & valList & " FROM [" & TempTable & "] " & _
                "LEFT JOIN [" & TargetTable & "] ON [" & TempTable & "].[" & KeyField & "] = [" & TargetTable & "].[" & KeyField & "] " & _
                "WHERE [" & TargetTable & "].[" & KeyField & "] IS NULL;"
    pDB.Execute sqlInsert, dbFailOnError
End Sub

' --- 以下、既存の演算メソッド (互換性のために維持) ---
Public Sub UpdateFieldFromField(src As String, tgt As String)
    Dim rs As DAO.Recordset: Set rs = pDB.OpenRecordset(TargetTable, dbOpenDynaset)
    Do Until rs.EOF
        rs.Edit: rs.Fields(tgt).value = Nz(rs.Fields(src).value, Null): rs.Update: rs.MoveNext
    Loop
    rs.Close
End Sub
' (省略: UpdateFieldFromDiff, UpdateFieldByConcat 等も同様に TargetTable を参照するよう修正して保持)

