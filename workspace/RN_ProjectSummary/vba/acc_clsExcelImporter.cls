VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsExcelImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --- 公開変数 (取込範囲指定用) ---
Public HeaderRow      As Long   ' タイトル行 (5など)
Public StartColumn    As Long   ' 開始列 (B列なら2など)
Public EndColumn      As Long   ' 終了列 (EK列なら141など)

' --- 内部保持変数 ---
Private pConn         As Object ' DAO.Database を保持
Private pTableName    As String
Private pMapping      As Object ' Dictionary (Excelタイトル -> Accessフィールド名)

' --- プロパティ ---
Public Property Let tableName(ByVal val As String):      pTableName = val: End Property
Public Property Get tableName() As String:               tableName = pTableName: End Property
Public Property Let TempTableName(ByVal val As String):  pTableName = val: End Property

'===========================================================
' メソッド : Init
'===========================================================
Public Sub Init(ByRef db As DAO.Database)
    Set pConn = db
    Set pMapping = CreateObject("Scripting.Dictionary")
End Sub

'===========================================================
' メソッド : AddMapping
'===========================================================
Public Sub AddMapping(ByVal excelTitle As String, ByVal accessFieldName As String)
    pMapping(excelTitle) = accessFieldName
End Sub

'===========================================================
' メソッド : ImportFromWorksheet (通常のリスト形式用)
'===========================================================
Public Sub ImportFromWorksheet(ByVal ws As Object)
    Dim dataArr     As Variant
    Dim lastRow     As Long
    Dim r As Long, c As Long
    Dim rsTemp      As DAO.Recordset
    
    lastRow = ws.Cells(ws.rows.count, StartColumn).End(-4162).Row ' -4162 = xlUp
    dataArr = ws.Range(ws.Cells(HeaderRow, StartColumn), ws.Cells(lastRow, EndColumn)).Value
    
    For c = 1 To UBound(dataArr, 2)
        If pMapping.Exists(CStr(Nz(dataArr(1, c), ""))) Then
            dataArr(1, c) = pMapping(CStr(dataArr(1, c)))
        Else
            dataArr(1, c) = "SKIP_COLUMN"
        End If
    Next c
    
    Set rsTemp = pConn.OpenRecordset(pTableName, 2)
    For r = 2 To UBound(dataArr, 1)
        rsTemp.AddNew
        For c = 1 To UBound(dataArr, 2)
            If CStr(dataArr(1, c)) <> "SKIP_COLUMN" Then
                On Error Resume Next
                rsTemp.Fields(CStr(dataArr(1, c))).Value = dataArr(r, c)
                On Error GoTo 0
            End If
        Next c
        rsTemp.Update
    Next r
    rsTemp.Close
End Sub

'===========================================================
' メソッド : ImportUnpivotedData (原本準拠：横に並んだ社員を縦に展開)
'===========================================================
Public Sub ImportUnpivotedData(ByVal lo As Object, ByVal worksiteName As String, ByVal filePath As String)
    Dim rs      As DAO.Recordset
    Dim dataArr As Variant
    Dim r As Long, c As Long
    Dim colName As String
    Dim colNorm As String
    
    ' 固定列のインデックスを取得（Normalize_Textを使用して表記揺れを吸収）
    Dim idxNo     As Long: idxNo = Get_ColumnIndex_Robust(lo, "No")
    Dim idxDate   As Long: idxDate = Get_ColumnIndex_Robust(lo, "年月")
    Dim idxCode   As Long: idxCode = Get_ColumnIndex_Robust(lo, "工事コード")
    Dim idxName   As Long: idxName = Get_ColumnIndex_Robust(lo, "工事名")
    Dim idxComm   As Long: idxComm = Get_ColumnIndex_Robust(lo, "コメント")
    
    ' テーブルを一括取得（高速化のため配列で処理）
    dataArr = lo.DataBodyRange.Value
    Set rs = pConn.OpenRecordset(pTableName, 2) ' 2 = dbOpenDynaset
    
    ' 行ループ
    For r = 1 To UBound(dataArr, 1)
        ' 列ループ (社員名が並んでいる列を走査)
        For c = 1 To lo.ListColumns.count
            colName = lo.ListColumns(c).Name
            colNorm = Normalize_Text(colName)
            
            ' 固定項目以外の列を「社員名」として扱う
            If colNorm <> "" And _
               colNorm <> Normalize_Text("No") And _
               colNorm <> Normalize_Text("年月") And _
               colNorm <> Normalize_Text("工事コード") And _
               colNorm <> Normalize_Text("工事名") And _
               colNorm <> Normalize_Text("コメント") Then
                
                ' 兼務率の値がある場合のみ、縦持ちレコードを作成
                If Nz(dataArr(r, c), 0) <> 0 Then
                    rs.AddNew
                    
                    ' ユニークIDの取得 (共通関数を使用)
                    rs!ImportID = Fetch_New_ImportID()
                    
                    ' 共通情報
                    rs!作業所名 = worksiteName
                    rs!元ファイルパス = filePath
                    
                    ' 行データからの転記
                    If idxDate > 0 Then rs!年月 = dataArr(r, idxDate)
                    If idxNo > 0 Then rs!No = dataArr(r, idxNo)
                    If idxCode > 0 Then rs!工事コード = dataArr(r, idxCode)
                    If idxName > 0 Then rs!工事名 = dataArr(r, idxName)
                    If idxComm > 0 Then rs!コメント = dataArr(r, idxComm)
                    
                    ' 列ヘッダーからの転記
                    rs!社員名 = colName
                    rs!兼務率割合 = dataArr(r, c)
                    
                    rs.Update
                End If
            End If
        Next c
    Next r
    
    rs.Close
End Sub



