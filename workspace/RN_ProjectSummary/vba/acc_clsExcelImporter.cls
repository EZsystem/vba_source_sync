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
' メソッド : ImportUnpivotedData (兼務率等の横持ち→縦持ち用)
'===========================================================
Public Sub ImportUnpivotedData(ByVal lo As Object, ByVal worksiteName As String, ByVal filePath As String)
    Dim rs      As DAO.Recordset
    Dim rowObj  As Object
    Dim col     As Object
    Dim baseID  As Long
    Dim rIdx    As Long
    
    ' UNIQUE ID生成ベース (ddHHmm)
    ' ※MMddHHmm(12312359)だと 12,312,359 * 1000 で Long(21億) を超えるリスクがあるため ddHHmm に短縮
    baseID = CLng(Format(Now, "ddHHmm"))
    
    Set rs = pConn.OpenRecordset(pTableName, 2)
    
    For Each rowObj In lo.ListRows
        rIdx = rIdx + 1
        For Each col In lo.ListColumns
            ' 列名が日付形式（年月）か判定
            If IsDate(col.Name) Then
                Dim val As Variant
                val = rowObj.Range(col.Index).Value
                
                ' 値がある場合のみ、縦持ちレコードとして追加
                If Nz(val, 0) <> 0 Then
                    rs.AddNew
                    ' ID: 時刻ベース + 3桁行番号 (21億の制限内)
                    rs!ImportID = (baseID * 1000) + (rIdx Mod 1000)
                    rs!作業所名 = worksiteName
                    rs!元ファイルパス = filePath
                    rs!年月 = col.Name
                    
                    ' 列名指定で基本項目を取得 (On Error Resume Next で存在しない列を回避)
                    On Error Resume Next
                    rs!No = rowObj.Range(lo.ListColumns("No").Index).Value
                    rs!工事コード = rowObj.Range(lo.ListColumns("工事コード").Index).Value
                    rs!工事名 = rowObj.Range(lo.ListColumns("工事名").Index).Value
                    rs!コメント = rowObj.Range(lo.ListColumns("コメント").Index).Value
                    rs!社員名 = rowObj.Range(lo.ListColumns("社員名").Index).Value
                    rs!兼務率割合 = val
                    On Error GoTo 0
                    
                    rs.Update
                End If
            End If
        Next col
    Next rowObj
    rs.Close
End Sub



