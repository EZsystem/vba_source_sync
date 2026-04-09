Attribute VB_Name = "xl_mod_TableUpdate"

'-------------------------------------
' Module: xl_mod_TableUpdate
' 説明  ：テーブル内データの体系的な更新（ID生成等）を管理
'-------------------------------------
Option Explicit

''' <summary>
''' tbl_内訳ID の 「内訳ID」のみを一括更新。
''' 仕様：内訳ID = 大分類ID + 大分類内の通し番号 (A01, A02...)
''' ※中分類IDなど、他の列は一切変更しません。
''' </summary>
Public Sub Update_UchiwakeIDs_All()
    Dim wsCat As Worksheet
    Dim tblID As ListObject, tblMasta As ListObject
    
    ' 1. 初期化・シート/テーブル取得
    On Error Resume Next
    Set wsCat = ThisWorkbook.Worksheets("分類")
    If wsCat Is Nothing Then
        MsgBox "シート「分類」が見つかりません。", vbCritical
        Exit Sub
    End If
    
    Set tblID = wsCat.ListObjects("tbl_内訳ID")
    Set tblMasta = wsCat.ListObjects("tbl_大分類")
    On Error GoTo 0
    
    If tblID Is Nothing Or tblMasta Is Nothing Then
        MsgBox "テーブル（tbl_内訳ID または tbl_大分類）が見つかりません。", vbCritical
        Exit Sub
    End If
    
    ' 高速化設定
    Call Optimize_Settings(True)
    
    ' 2. 大分類マスタの辞書作成 (Key:大分類名, Value:大分類ID)
    Dim dictMasta As Object: Set dictMasta = CreateObject("Scripting.Dictionary")
    Dim mastaData As Variant: mastaData = tblMasta.DataBodyRange.Value
    Dim i As Long
    For i = 1 To UBound(mastaData, 1)
        dictMasta(Trim(CStr(mastaData(i, 2)))) = Trim(CStr(mastaData(i, 1)))
    Next i
    
    ' 3. ID生成ロジック：大分類内でのカウント管理
    Dim colLarge As Long: colLarge = tblID.ListColumns("大分類").Index
    Dim colUchiID As Long: colUchiID = tblID.ListColumns("内訳ID").Index
    
    Dim dataArr As Variant: dataArr = tblID.DataBodyRange.Value
    Dim resUchiArr() As Variant: ReDim resUchiArr(1 To UBound(dataArr, 1), 1 To 1)
    
    ' 大分類ごとの「通し番号」カウンタ（内訳ID用）
    Dim dictCounterLarge As Object: Set dictCounterLarge = CreateObject("Scripting.Dictionary")
    
    Dim largeName As String, prefix As String
    
    For i = 1 To UBound(dataArr, 1)
        largeName = Trim(CStr(dataArr(i, colLarge)))
        
        ' 接頭辞（大分類ID）を取得
        If dictMasta.Exists(largeName) Then
            prefix = dictMasta(largeName)
        Else
            prefix = "?"
        End If
        
        ' --- 内訳ID 生成：指定された大分類内での単純通し番号 ---
        dictCounterLarge(largeName) = dictCounterLarge(largeName) + 1
        resUchiArr(i, 1) = prefix & Format(dictCounterLarge(largeName), "00")
    Next i
    
    ' 4. テーブルへ一括書き込み（内訳ID列のみ）
    tblID.ListColumns(colUchiID).DataBodyRange.Value = resUchiArr
    
    ' 終了処理
    Call Optimize_Settings(False)
    MsgBox "内訳IDの一括更新が完了しました。", vbInformation
    
End Sub

''' <summary>
''' 「出力範囲→」シートより右側のシートをすべて1つのPDFとして一括出力
''' </summary>
Public Sub Export_Sheets_To_PDF()
    Dim wsStart As Worksheet
    Dim startIndex As Integer, i As Integer
    Dim sheetNames() As String
    Dim count As Integer
    Dim originalSheet As Object
    Dim savePath As Variant
    
    ' 1. 「出力範囲→」シートの特定
    On Error Resume Next
    Set wsStart = ThisWorkbook.Worksheets("出力範囲→")
    On Error GoTo 0
    
    If wsStart Is Nothing Then
        MsgBox "シート「出力範囲→」が見つかりません。" & vbCrLf & _
               "このシートより右側のシートが出力対象となります。", vbCritical
        Exit Sub
    End If
    
    startIndex = wsStart.Index
    
    ' 2. 対象シート名の収集（「出力範囲→」より右側の全シート）
    count = ThisWorkbook.Worksheets.count - startIndex
    If count <= 0 Then
        MsgBox "出力対象のシート（「出力範囲→」より右側）がありません。", vbExclamation
        Exit Sub
    End If
    
    ReDim sheetNames(0 To count - 1)
    For i = 1 To count
        sheetNames(i - 1) = ThisWorkbook.Worksheets(startIndex + i).Name
    Next i
    
    ' 3. 保存先の指定（ファイル名ダイアログ）
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="【PDF出力】長期修繕計画_" & Format(Now, "yyyymmdd"), _
        FileFilter:="PDFファイル (*.pdf), *.pdf", _
        Title:="PDF保存先の指定")
    
    If savePath = False Then Exit Sub ' キャンセル時
    
    ' 4. 出力実行
    Set originalSheet = ActiveSheet
    
    ' 高速化設定（画面更新停止のみ）
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    ' 指定したシートを一括選択してエクスポート
    Sheets(sheetNames).Select
    
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    originalSheet.Select
    Application.ScreenUpdating = True
    MsgBox "PDFの一括出力が完了しました。", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Not originalSheet Is Nothing Then originalSheet.Select
    MsgBox "PDF出力中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
End Sub

''' <summary>
''' 画面更新等の高速化設定一括切替
''' </summary>

Private Sub Optimize_Settings(isStart As Boolean)
    With Application
        .ScreenUpdating = Not isStart
        .Calculation = IIf(isStart, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not isStart
    End With
End Sub
