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

' --- 内部保持変数 ---
Private pConn         As Object ' DAO.Database を保持
Private pMapping      As Object ' Dictionary (Excelタイトル -> Accessフィールド名)

' --- 呼び出し側から渡される設定値 (実数は持たない) ---
Public TableName      As String
Public HeaderRow      As Long   ' タイトル行 (5など)
Public StartColumn    As Long   ' 開始列 (B列なら2など)
Public EndColumn      As Long   ' 終了列 (EK列なら141など)

'===========================================================
' メソッド : Init
'===========================================================
Public Sub Init(ByRef Db As DAO.Database)
    Set pConn = Db
    Set pMapping = CreateObject("Scripting.Dictionary")
End Sub

'===========================================================
' メソッド : AddMapping
' 説明：Excelの元の名前と、Accessでの名前のペアを登録する
'===========================================================
Public Sub AddMapping(ByVal excelTitle As String, ByVal accessFieldName As String)
    pMapping(excelTitle) = accessFieldName
End Sub

'===========================================================
' メソッド : ImportFromWorksheet
' 説明：物理置換ロジックを用いて仮テーブルへ書き込む
'===========================================================
Public Sub ImportFromWorksheet(ByVal wS As Object)
    Dim dataArr     As Variant
    Dim lastRow     As Long
    Dim r As Long, c As Long
    Dim rsTemp      As DAO.Recordset
    
    ' 1. Excelを配列に一括読込 (プロパティで指定された範囲を使用)
    lastRow = wS.Cells(wS.Rows.Count, StartColumn).End(-4162).row ' -4162 = xlUp
    dataArr = wS.Range(wS.Cells(HeaderRow, StartColumn), wS.Cells(lastRow, EndColumn)).Value
    
    ' 2. 配列1行目(タイトル)を物理的に翻訳(置換)
    For c = 1 To UBound(dataArr, 2)
        Dim currentTitle As String
        currentTitle = CStr(Nz(dataArr(1, c), ""))
        
        If pMapping.Exists(currentTitle) Then
            ' 配列の値を、マスタから登録されたAccess列名で上書き
            dataArr(1, c) = pMapping(currentTitle)
        Else
            dataArr(1, c) = "SKIP_COLUMN"
        End If
    Next c
    
    ' 3. テーブルへの流し込み
    Set rsTemp = pConn.OpenRecordset(TableName, 2) ' 2 = dbOpenDynaset
    
    For r = 2 To UBound(dataArr, 1) ' 2行目からデータ開始
        rsTemp.AddNew
        For c = 1 To UBound(dataArr, 2)
            Dim targetField As String
            targetField = CStr(dataArr(1, c))
            
            If targetField <> "SKIP_COLUMN" Then
                On Error Resume Next
                rsTemp.fields(targetField).Value = dataArr(r, c)
                On Error GoTo 0
            End If
        Next c
        rsTemp.Update
    Next r
    rsTemp.Close
End Sub



