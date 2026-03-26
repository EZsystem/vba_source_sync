VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsExcelImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class: acc_clsExcelImporter
' 説明 : Excel配列データを仮テーブルへ取り込み、メインへ反映する
'=====================================
Option Compare Database
Option Explicit

Private pDB As DAO.Database
Private pDictMapping As Object
Private pDictValidFields As Object
Private pColMapping As Object

Public TableName As String
Public TempTableName As String
Public MappingTableName As String
Public ImportColumnMax As Long

' 初期化
Public Sub Init()
    Set pDB = CurrentDb
    Set pDictMapping = CreateObject("Scripting.Dictionary")
    Set pDictValidFields = CreateObject("Scripting.Dictionary")
    Set pColMapping = CreateObject("Scripting.Dictionary")
End Sub

' マッピング情報の準備
Public Sub PrepareMappings()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim rsMap As DAO.Recordset
    Dim colNum As Variant
    Dim fieldName As String
    Dim defVal As Variant
    Dim i As Long

    ' 1. 出先テーブルの有効フィールドを把握
    Set tdf = pDB.TableDefs(TableName)
    For Each fld In tdf.Fields
        pDictValidFields.Add fld.Name, True
    Next fld

    ' 2. マッピングテーブルの読み込み
    Set rsMap = pDB.OpenRecordset("SELECT * FROM [" & MappingTableName & "]")
    Do While Not rsMap.EOF
        colNum = rsMap("列番号").value
        fieldName = rsMap("タイトル名").value
        defVal = rsMap("デフォルト値").value
        If Not pDictMapping.Exists(colNum) Then
            pDictMapping.Add colNum, Array(fieldName, defVal)
        End If
        rsMap.MoveNext
    Loop
    rsMap.Close

    ' 3. 列番号とフィールド名の対応付け
    For i = 1 To ImportColumnMax
        If pDictMapping.Exists(i) Then
            If pDictValidFields.Exists(pDictMapping(i)(0)) Then
                pColMapping.Add i, pDictMapping(i)(0)
            End If
        End If
    Next i
End Sub ' ← PrepareMappings 終了

' 配列からのインポート
Public Sub Import_From_Array(ByVal dataArr As Variant, Optional ByVal basicCode As Variant)
    Dim rs As DAO.Recordset
    Dim r As Long
    Dim c As Variant
    
    Set rs = pDB.OpenRecordset(TempTableName, dbOpenDynaset)
    
    For r = 1 To UBound(dataArr, 1)
        ' A列が2桁の数値（行番号）である場合のみ取り込む
        If Not (IsNumeric(dataArr(r, 1)) And Len(CStr(dataArr(r, 1))) = 2) Then GoTo SkipRow
        
        rs.AddNew
        ' 基本工事名の補完（A2セルの値など）
        If Not IsMissing(basicCode) Then
            If FieldExists(rs, "基本工事名") Then rs.Fields("基本工事名").value = basicCode
        End If

        ' マッピングに基づく列の流し込み
        For Each c In pColMapping.Keys
            Dim val As Variant: val = dataArr(r, c)
            If IsNull(val) Or Trim(CStr(val)) = "" Then
                rs.Fields(pColMapping(c)).value = pDictMapping(c)(1) ' デフォルト値
            Else
                rs.Fields(pColMapping(c)).value = val
            End If
        Next c
        rs.Update
SkipRow:
    Next r
    rs.Close
End Sub ' ← Import_From_Array 終了

' メインテーブルへ追加
Public Sub Commit_To_Main()
    pDB.Execute "INSERT INTO [" & TableName & "] SELECT * FROM [" & TempTableName & "]", dbFailOnError
End Sub ' ← Commit_To_Main 終了
