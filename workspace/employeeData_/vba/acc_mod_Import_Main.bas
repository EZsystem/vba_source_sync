Attribute VB_Name = "acc_mod_Import_Main"
'-------------------------------------
' Module: acc_mod_Import_Main (EZsystem Refactored)
' 修正内容：
'    1. ファイルパス（sourcePath）を受け取り、テーブルへ保存する機能を追加
'    2. 引数の数（シグネチャ）を呼び出し元と一致するよう修正
'-------------------------------------
Option Explicit

Private Const TARGET_TABLE As String = "at_kenmuTemp"
Private Const TARGET_SHEET As String = "職員兼務率"
Private Const TARGET_LISTOBJ As String = "xt_kenmu"
Private Const INPUT_FOLDER As String = "D:\My_Projects\RN管理表関係\職員兼務率\inputdata\"

'--------------------------------------------
' プロシージャ名： Run_Kenmu_Import_EZ
'--------------------------------------------
Public Sub Run_Kenmu_Import_EZ()
    Dim importer As New acc_clsExcelImporter
    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim fileName As String
    Dim filePath As String
    Dim fileCount As Long
    
    Set xlApp = CreateObject("Excel.Application")
    Call Fast_Mode_Toggle(True, xlApp)
    
    On Error GoTo ErrLine
    
    CurrentDb.Execute "DELETE * FROM [" & TARGET_TABLE & "];", dbFailOnError

    importer.Init
    importer.TempTableName = TARGET_TABLE

    fileName = Dir(INPUT_FOLDER & "*.xls*")
    
    If fileName = "" Then
        MsgBox "指定されたフォルダにExcelファイルが見つかりません:" & vbCrLf & INPUT_FOLDER, vbExclamation
        GoTo CleanUp
    End If

    Do While fileName <> ""
        filePath = INPUT_FOLDER & fileName
        fileCount = fileCount + 1
        
        Set wb = xlApp.Workbooks.Open(filePath, ReadOnly:=True)
        
        On Error Resume Next
        Set ws = wb.Worksheets(TARGET_SHEET)
        On Error GoTo ErrLine
        
        If Not ws Is Nothing Then
            Dim lo As Object
            Set lo = ws.ListObjects(TARGET_LISTOBJ)
            
            Dim worksiteName As String
            worksiteName = ws.Range("D2").value
            
            If Not lo Is Nothing Then
                ' ★第3引数として filePath を渡します
                Call Process_Kenmu_Data_Custom(lo, worksiteName, filePath)
            End If
        End If
        
        wb.Close SaveChanges:=False
        Set ws = Nothing
        fileName = Dir()
    Loop

    Call Notify_Smart_Popup(fileCount & " 件のファイルをインポートしました。", "完了通知")

CleanUp:
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    Call Fast_Mode_Toggle(False)
    Exit Sub

ErrLine:
    MsgBox "エラー発生 (" & fileName & "): " & Err.Description, vbCritical
    Resume CleanUp
End Sub

'--------------------------------------------
' 内部補助：兼務率特有のデータ変換ロジック
' ★修正ポイント：引数に ByVal sourcePath As String を追加
'--------------------------------------------
Public Sub Process_Kenmu_Data_Custom(ByRef lo As Object, ByVal worksiteName As String, ByVal sourcePath As String)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim dataArr As Variant
    Dim r As Long, c As Long
    Dim colName As String
    Dim normCol As String
    
    Dim idxNo As Long:    idxNo = Get_ColumnIndex_Robust(lo, "No")
    Dim idxYear As Long:  idxYear = Get_ColumnIndex_Robust(lo, "年月")
    Dim idxCode As Long:  idxCode = Get_ColumnIndex_Robust(lo, "工事コード")
    Dim idxName As Long:  idxName = Get_ColumnIndex_Robust(lo, "工事名")
    Dim idxComm As Long:  idxComm = Get_ColumnIndex_Robust(lo, "コメント")
    
    dataArr = lo.DataBodyRange.value
    Set rs = db.OpenRecordset(TARGET_TABLE, dbOpenDynaset)
    
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To lo.ListColumns.Count
            colName = lo.ListColumns(c).Name
            normCol = Normalize_Text(colName)
            
            If InStr(colName, vbLf) = 0 And _
               UCase(normCol) <> "NO" And _
               normCol <> Normalize_Text("年月") And _
               normCol <> Normalize_Text("工事コード") And _
               normCol <> Normalize_Text("工事名") And _
               normCol <> Normalize_Text("コメント") Then
               
                If Not IsEmpty(dataArr(r, c)) And dataArr(r, c) <> 0 Then
                    rs.AddNew
                    
                    ' ★追加：元ファイルパスをテーブルのフィールドへ保存
                    rs!元ファイルパス = sourcePath
                    rs!作業所名 = worksiteName
                    
                    If idxNo > 0 Then rs!No = dataArr(r, idxNo)
                    If idxYear > 0 Then rs!年月 = dataArr(r, idxYear)
                    If idxCode > 0 Then rs!工事コード = dataArr(r, idxCode)
                    If idxName > 0 Then rs!工事名 = dataArr(r, idxName)
                    If idxComm > 0 Then rs!コメント = dataArr(r, idxComm)
                    rs!社員名 = colName
                    rs!兼務率割合 = dataArr(r, c)
                    rs.Update
                End If
            End If
        Next c
    Next r
    rs.Close
End Sub

