Attribute VB_Name = "mod_MainControl"
'-------------------------------------
' Module: mod_MainControl
' 説明　：クラスモジュールを使用した統合実行制御
'-------------------------------------
Option Explicit

' ==========================================
'  1. データ紐付け・更新系 (cls_DataLookup利用)
' ==========================================

' 計画表のH列（小分類）を更新
' 元ファイル: she05_Subcategory
Public Sub Main_Update_Subcategory()
    Dim logic As New cls_DataLookup
    
    ' 1. 内訳テーブルを読み込み、IDと小分類を紐付け
    logic.Initialize tbl:=ThisWorkbook.Worksheets("内訳").ListObjects("tbl_内訳"), _
                     keyColName:="内訳ID", _
                     valColName:="小分類"
    
    ' 2. 計画表のG列(Key)を見て、H列(Out)に書き出し (9行目から)
    logic.ExecuteBatchUpdate targetWs:=ThisWorkbook.Worksheets("計画表"), _
                             keyColLetter:="G", _
                             outColLetter:="H", _
                             startRow:=9
                             
    MsgBox "小分類の更新が完了しました", vbInformation
End Sub

' 計画表のI列（修繕内容）を更新
' 元ファイル: she05_Repair
Public Sub Main_Update_RepairContent()
    Dim logic As New cls_DataLookup
    
    ' 1. 内訳テーブルを読み込み、IDと修繕内容を紐付け
    logic.Initialize tbl:=ThisWorkbook.Worksheets("内訳").ListObjects("tbl_内訳"), _
                     keyColName:="内訳ID", _
                     valColName:="修繕内容"
    
    ' 2. 計画表のG列(Key)を見て、I列(Out)に書き出し (15行目から)
    logic.ExecuteBatchUpdate targetWs:=ThisWorkbook.Worksheets("計画表"), _
                             keyColLetter:="G", _
                             outColLetter:="I", _
                             startRow:=15
                             
    MsgBox "修繕内容の更新が完了しました", vbInformation
End Sub


' ==========================================
'  2. シート生成・削除系 (cls_SheetManager利用)
' ==========================================

' テンプレート1(出力内訳og)から生成
' 元ファイル: she04_1Generator
Public Sub Main_Generate_FromTemplate1()
    Dim mgr As New cls_SheetManager
    
    ' IDリストをセット
    mgr.Initialize tbl:=ThisWorkbook.Worksheets("分類").ListObjects("tbl_内訳ID")
    
    ' 生成実行
    mgr.CreateSheets templateName:="出力内訳og"
End Sub

' テンプレート2(出力内訳og2)から生成
' 元ファイル: Module4 (Execute_GenerateSheetsByID2)
Public Sub Main_Generate_FromTemplate2()
    Dim mgr As New cls_SheetManager
    
    ' IDリストをセット
    mgr.Initialize tbl:=ThisWorkbook.Worksheets("分類").ListObjects("tbl_内訳ID")
    
    ' 生成実行
    mgr.CreateSheets templateName:="出力内訳og2"
End Sub

' シート一括削除
' 元ファイル: she04_2Deleter
Public Sub Main_Delete_Sheets()
    Dim mgr As New cls_SheetManager
    
    mgr.Initialize tbl:=ThisWorkbook.Worksheets("分類").ListObjects("tbl_内訳ID")
    
    ' 削除実行
    mgr.DeleteSheets
End Sub


' ==========================================
'  3. 印刷設定系 (cls_PrintSetting利用)
' ==========================================

' 印刷範囲設定 (パターン1: 24行/23行)
' 元ファイル: she04_5PrintRan
Public Sub Main_SetPrintArea_Pattern1()
    Call Execute_PrintSetting(24, 23)
End Sub

' 印刷範囲設定 (パターン2: 41行/46行)
' 元ファイル: she04_5PrintRan2
Public Sub Main_SetPrintArea_Pattern2()
    Call Execute_PrintSetting(41, 46)
End Sub

' 共通実行用プロシージャ（Private）
Private Sub Execute_PrintSetting(firstRows As Long, nextRows As Long)
    Dim printer As New cls_PrintSetting
    Dim tbl As ListObject
    Dim idRange As Range, cell As Range
    Dim wsName As String, targetWs As Worksheet
    Dim count As Long
    
    ' 設定値をクラスに渡す
    printer.Configure firstRows, nextRows
    
    Set tbl = ThisWorkbook.Worksheets("分類").ListObjects("tbl_内訳ID")
    Set idRange = tbl.ListColumns("内訳ID").DataBodyRange
    
    Application.ScreenUpdating = False
    
    For Each cell In idRange
        wsName = Trim(CStr(cell.value))
        If wsName <> "" Then
            On Error Resume Next
            Set targetWs = ThisWorkbook.Worksheets(wsName)
            On Error GoTo 0
            
            If Not targetWs Is Nothing Then
                If printer.SetPrintArea(targetWs) Then
                    count = count + 1
                End If
                Set targetWs = Nothing
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "印刷範囲設定完了: " & count & "件", vbInformation
End Sub

' フッター設定処理
' 元ファイル: she04_7FooterFor
Public Sub Main_SetFooter_All()
    Dim printer As New cls_PrintSetting
    Dim tbl As ListObject
    Dim i As Long
    Dim wsName As String, major As String, minor As String
    Dim startPage As Long, totalPage As String
    Dim ws As Worksheet
    
    Set tbl = ThisWorkbook.Worksheets("分類").ListObjects("tbl_内訳ID")
    totalPage = CStr(ThisWorkbook.Worksheets("分類").Range("M5").value)
    
    Application.ScreenUpdating = False
    
    With tbl
        For i = 1 To .DataBodyRange.Rows.count
            ' テーブルから値取得
            wsName = .ListColumns("内訳ID").DataBodyRange(i).value
            major = .ListColumns("大分類").DataBodyRange(i).value
            minor = .ListColumns("中分類").DataBodyRange(i).value
            startPage = .ListColumns("累計ページ").DataBodyRange(i).value
            
            If wsName <> "" Then
                On Error Resume Next
                Set ws = ThisWorkbook.Worksheets(wsName)
                On Error GoTo 0
                
                If Not ws Is Nothing Then
                    ' フッター設定実行
                    printer.SetFooter ws, _
                                      wsName, _
                                      major & "：" & minor, _
                                      "P&P/" & totalPage, _
                                      startPage
                End If
                Set ws = Nothing
            End If
        Next i
    End With
    
    Application.ScreenUpdating = True
    MsgBox "フッター設定完了", vbInformation
End Sub

' ページ数カウント処理
' 元ファイル: she04_6PageCount
Public Sub Main_Update_PageCounts()
    Dim printer As New cls_PrintSetting
    Dim tbl As ListObject
    Dim i As Long, wsName As String
    Dim ws As Worksheet
    Dim pCount As Long
    
    Set tbl = ThisWorkbook.Worksheets("分類").ListObjects("tbl_内訳ID")
    
    Application.ScreenUpdating = False
    
    For i = 1 To tbl.DataBodyRange.Rows.count
        wsName = tbl.ListColumns("内訳ID").DataBodyRange(i).value
        pCount = 0
        
        If wsName <> "" Then
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(wsName)
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                pCount = printer.GetActualPageCount(ws)
            End If
            Set ws = Nothing
        End If
        
        ' 結果を「単ページ数」列に書き込み
        tbl.ListColumns("単ページ数").DataBodyRange(i).value = pCount
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "ページ数集計完了", vbInformation
End Sub


' ==========================================
'  4. ユーティリティ系 (cls_CellUtility利用)
' ==========================================

' 選択範囲を半角にする
Public Sub Util_ToNarrow()
    Dim util As New cls_CellUtility
    If TypeName(Selection) = "Range" Then
        util.ConvertToNarrow Selection
        MsgBox "半角変換完了", vbInformation
    End If
End Sub

' 選択範囲を全角にする（数値はスキップ）
Public Sub Util_ToWide()
    Dim util As New cls_CellUtility
    If TypeName(Selection) = "Range" Then
        util.ConvertToWide Selection, True
        MsgBox "全角変換完了", vbInformation
    End If
End Sub

' ○の上のセルに○をつける
Public Sub Util_MarkCircleAbove()
    Dim util As New cls_CellUtility
    Dim cnt As Long
    If TypeName(Selection) = "Range" Then
        cnt = util.MarkCircleAbove(Selection)
        MsgBox cnt & "件の○を記入しました", vbInformation
    End If
End Sub


' ==========================================
'  5. データ転記・抽出系 (cls_DataLookup利用)
' ==========================================

' 分類テーブルの基本情報を各シート(B8, C8等)へ転記
' 元ファイル: she04_3DataTransfer
Public Sub Main_Transfer_HeaderData()
    Dim tbl As ListObject
    Dim i As Long
    Dim wsName As String
    Dim ws As Worksheet
    
    Set tbl = ThisWorkbook.Worksheets("分類").ListObjects("tbl_内訳ID")
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To tbl.DataBodyRange.Rows.count
        ' テーブル値取得
        wsName = tbl.ListColumns("内訳ID").DataBodyRange(i).value
        
        If wsName <> "" Then
            Set ws = Nothing
            Set ws = ThisWorkbook.Worksheets(wsName)
            
            If Not ws Is Nothing Then
                ' 転記実行
                ws.Range("B8").value = tbl.ListColumns("大分類").DataBodyRange(i).value
                ws.Range("C8").value = tbl.ListColumns("中分類").DataBodyRange(i).value
                ws.Range("C9").value = tbl.ListColumns("種類").DataBodyRange(i).value
                ws.Range("C7").value = wsName ' 内訳ID
                ws.Range("C10").value = tbl.ListColumns("更新周期").DataBodyRange(i).value
            End If
        End If
    Next i
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    MsgBox "基本情報の転記が完了しました", vbInformation
End Sub

' 内訳データを抽出して各シート(13行目〜)へ転記
' 元ファイル: she04_0Breakdown / she04_4BreakdownGo
' ★修正：タイトル行を自動特定して抽出実行
Public Sub Main_Extract_BreakdownData()
    Dim lookup As New cls_DataLookup
    Dim tblID As ListObject, tblData As ListObject
    Dim i As Long
    Dim wsName As String
    Dim ws As Worksheet
    Dim headerRow As Long
    
    Set tblID = ThisWorkbook.Worksheets("分類").ListObjects("tbl_内訳ID")
    Set tblData = ThisWorkbook.Worksheets("内訳").ListObjects("tbl_内訳")
    
    Application.ScreenUpdating = False
    
    For i = 1 To tblID.DataBodyRange.Rows.count
        wsName = tblID.ListColumns("内訳ID").DataBodyRange(i).value
        
        If wsName <> "" Then
            Set ws = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(wsName)
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                ' テーブル情報からヘッダー行を自動特定
                headerRow = GetHeaderRow(ws)
                
                If headerRow > 0 Then
                    ' (ソーステーブル, 対象シート, 検索ID, 出力開始行, タイトル行)
                    lookup.ExecuteDataExtraction tblData, ws, wsName, 13, headerRow
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "内訳データの抽出・転記が完了しました", vbInformation
End Sub

' 各シートのL9セルの値を分類テーブルへ書き戻す
' 元ファイル: she04_11TableSum
Public Sub Main_Sync_L9_To_Table()
    Dim tbl As ListObject
    Dim i As Long
    Dim wsName As String
    Dim ws As Worksheet
    Dim val As Variant
    
    Set tbl = ThisWorkbook.Worksheets("分類").ListObjects("tbl_内訳ID")
    
    Application.ScreenUpdating = False
    
    For i = 1 To tbl.DataBodyRange.Rows.count
        wsName = tbl.ListColumns("内訳ID").DataBodyRange(i).value
        
        If wsName <> "" Then
            Set ws = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(wsName)
            val = ws.Range("L9").value
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                tbl.ListColumns("出力シート集計").DataBodyRange(i).value = val
            Else
                tbl.ListColumns("出力シート集計").DataBodyRange(i).value = "シートなし"
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "集計値(L9)の同期が完了しました", vbInformation
End Sub


' ==========================================
'  6. データ整合性・ロジック系 (cls_DataLogic利用)
' ==========================================

' 内訳IDマッピング (Single)
' 元機能: she02_IdMappingSingle
Public Sub Main_Logic_IdMapping_Single()
    Dim logic As New cls_DataLogic
    Dim wsBreakdown As Worksheet
    Dim wsCategory As Worksheet
    
    Set wsBreakdown = ThisWorkbook.Worksheets("内訳")
    Set wsCategory = ThisWorkbook.Worksheets("分類")
    
    Application.ScreenUpdating = False
    
    logic.MapBreakdownIDs wsBreakdown, wsCategory
    
    Application.ScreenUpdating = True
    MsgBox "内訳IDマッピング処理が完了しました", vbInformation
End Sub

' 内訳IDマッピング (Double: Single + 上行コピー)
' 修正：cls_DataLogic (旧) から cls_IdMappingProcessor (新・高速版) へ変更
' 元機能: she02_IdMappingDouble
Public Sub Main_Logic_IdMapping_Double()
    Dim processor As New cls_IdMappingProcessor
    
    ' クラス内で ScreenUpdating = False 等の高速化設定も行っているため、
    ' ここでの ScreenUpdating 制御は削除しても良いですが、
    ' 全体の安全のため入れておいても問題ありません。
    Application.ScreenUpdating = False
    
    ' マッピングとIdpa補正を一括実行
    processor.ExecuteFullProcess
    
    Application.ScreenUpdating = True
    
    ' メッセージはクラス内でも出していますが、重複する場合は
    ' クラス側のMsgBoxを削除するか、こちらのMsgBoxを削除してください。
    ' （基本的にはクラス側に任せるか、クラス側をサイレントにしてここで出すのが作法です）
End Sub

' 重複チェック
' 元機能: she02_DupliCheck
Public Sub Main_Logic_CheckDuplicates()
    Dim logic As New cls_DataLogic
    Dim wsBreakdown As Worksheet
    
    Set wsBreakdown = ThisWorkbook.Worksheets("内訳")
    
    Application.ScreenUpdating = False
    
    logic.CheckDuplicates wsBreakdown
    
    Application.ScreenUpdating = True
    MsgBox "重複チェック処理が完了しました", vbInformation
End Sub

' BELCAコード最大値更新
' 元機能: she02_BELCAMax
Public Sub Main_Logic_UpdateBelcaMax()
    Dim logic As New cls_DataLogic
    Dim wsBreakdown As Worksheet
    Dim wsBelca As Worksheet
    
    Set wsBreakdown = ThisWorkbook.Worksheets("内訳")
    Set wsBelca = ThisWorkbook.Worksheets("BELCA")
    
    Application.ScreenUpdating = False
    
    logic.UpdateBelcaMax wsBreakdown, wsBelca
    
    Application.ScreenUpdating = True
    MsgBox "BELCAMax列の更新が完了しました", vbInformation
End Sub


' ==========================================
'  7. 個別シート操作系 (アクティブシート対象)
' ==========================================

' アクティブシートのC7セルのIDを元に、データ抽出を実行
' 元機能: she04_0Breakdown.ExecuteBreakdownDataExtraction
' ★修正：タイトル行を自動特定して抽出実行
Public Sub Main_Extract_ActiveSheet()
    Dim lookup As New cls_DataLookup
    Dim tblData As ListObject
    Dim targetId As String
    Dim ws As Worksheet
    Dim headerRow As Long
    
    Set ws = ActiveSheet
    
    ' ID取得 (C7セル)
    targetId = Trim(CStr(ws.Range("C7").value))
    If targetId = "" Then
        MsgBox "C7セルに内訳IDが入力されていません", vbExclamation
        Exit Sub
    End If
    
    ' 内訳テーブル取得
    On Error Resume Next
    Set tblData = ThisWorkbook.Worksheets("内訳").ListObjects("tbl_内訳")
    On Error GoTo 0
    
    If tblData Is Nothing Then
        MsgBox "内訳シートまたはtbl_内訳が見つかりません", vbCritical
        Exit Sub
    End If
    
    ' テーブル情報からヘッダー行を自動特定
    headerRow = GetHeaderRow(ws)
    
    If headerRow = 0 Then
        MsgBox "タイトル行を特定できませんでした（標準:6行目）。" & vbCrLf & _
               "テーブル設定などを確認してください。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 抽出実行 (開始行13、タイトル行headerRow)
    lookup.ExecuteDataExtraction tblData, ws, targetId, 13, headerRow
    
    Application.ScreenUpdating = True
    MsgBox "抽出完了: " & targetId, vbInformation
End Sub

' ------------------------------------------------------------------
' [内部ヘルパー] シート情報から見出し行番号を取得する
' ------------------------------------------------------------------
Private Function GetHeaderRow(ws As Worksheet) As Long
    ' パターン1: シートに「テーブル(ListObject)」がある場合
    If ws.ListObjects.count > 0 Then
        ' 1つ目のテーブルの見出し行を返す
        GetHeaderRow = ws.ListObjects(1).HeaderRowRange.Row
        Exit Function
    End If
    
    ' パターン2: テーブル機能を使っていない場合（デフォルト6行目）
    GetHeaderRow = 6
End Function


' ==========================================
'  フォーム起動用マクロ
' ==========================================
Public Sub Show_Launcher()
    ' vbModeless を指定すると、フォームを開いたままシート操作ができます
    frm_Launcher.Show vbModeless
End Sub

