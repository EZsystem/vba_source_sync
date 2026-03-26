Attribute VB_Name = "mod_IcubeImport_Refined"
'Attribute VB_Name = "mod_IcubeImport_Refined"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名   : mod_IcubeImport_Refined
' 概要           : iCube工事決定累積データ 統合インポート制御モジュール
' 依存コンポーネント:
'   - クラス     : acc_clsExcelImporter (Excel物理置換インポート)
'   - クラス     : acc_clsTableTransfer (テーブル間安全転写)
'   - 定数管理   : acc_mod_MappingTemplate
'===================================================================================================

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Run_IcubeImport_FullStep
'---------------------------------------------------------------------------------------------------
Public Sub Run_IcubeImport_FullStep()
    ' --- 1. 設定値定義 ---
    ' 共通定数(acc_mod_MappingTemplate)を使用するように修正
    Dim TBL_TEMP      As String: TBL_TEMP = AT_ICUBE_IMPORT_WORK
    Dim TBL_TARGET    As String: TBL_TARGET = AT_ICUBE
    
    Const EXCEL_HEADER  As Long = 5      ' ヘッダー行番号
    Const EXCEL_START_C As Long = 2      ' 取込開始列（B列）
    Const EXCEL_END_C   As Long = 141    ' 取込終了列（EK列）
    Const KEY_FIELD     As String = "No" ' 重複排除および更新時のキーフィールド

    Dim db          As DAO.Database: Set db = CurrentDb
    Dim wsTrans     As DAO.Workspace: Set wsTrans = DBEngine.Workspaces(0)
    
    ' インスタンス生成
    Dim pImporter   As New acc_clsExcelImporter
    Dim pTransfer   As New acc_clsTableTransfer
    
    Dim xlApp       As Object
    Dim wb          As Object
    Dim filePath    As String
    Dim RsMaster    As DAO.Recordset
    Dim isInTrans   As Boolean: isInTrans = False
    
    On Error GoTo Err_Handler

    ' --- 2. 外部ファイル選択 ---
    filePath = Pick_File()
    If filePath = "" Then Exit Sub

    ' --- 3. クラスの初期化とインポート設定 ---
    pImporter.Init db
    pImporter.TableName = TBL_TEMP
    pImporter.HeaderRow = EXCEL_HEADER
    pImporter.StartColumn = EXCEL_START_C
    pImporter.EndColumn = EXCEL_END_C
    
    pTransfer.Init db

    ' --- 4. マッピングルールのロード ---
    ' テーブル名を定数 AT_ICUBE_COL_SETTING に差し替え、角括弧で保護
    Set RsMaster = db.OpenRecordset("SELECT [タイトル名_デフォルト], [タイトル名_置換え後] FROM [" & AT_ICUBE_COL_SETTING & "] WHERE Nz(取込フラグ, False) = True")
    Do Until RsMaster.EOF
        pImporter.AddMapping CStr(RsMaster![タイトル名_デフォルト]), CStr(RsMaster![タイトル名_置換え後])
        RsMaster.MoveNext
    Loop
    RsMaster.Close

    ' --- 5. トランザクション開始 ---
    wsTrans.BeginTrans
    isInTrans = True

    ' --- 6. 仮テーブルのクリア ---
    db.Execute "DELETE FROM [" & TBL_TEMP & "];", 128 ' 128 = dbFailOnError
    
    ' --- 7. Excelアプリケーション制御 ---
    Set xlApp = CreateObject("Excel.Application")
    Set wb = xlApp.Workbooks.Open(filePath, ReadOnly:=True)
    
    ' インポート実行
    pImporter.ImportFromWorksheet wb.Sheets(1)
    
    wb.Close False
    xlApp.Quit
    Set xlApp = Nothing

    ' --- 8. 本番テーブル(at_Icube)への型変換転写 ---
    ' TBL_TARGET に "at_Icube" が渡されるため TableDefs エラーが解消されます
    pTransfer.ExecuteTransfer TBL_TEMP, TBL_TARGET, KEY_FIELD

    ' --- 9. 処理の確定 ---
    wsTrans.CommitTrans
    isInTrans = False

    MsgBox "クラスによる統合インポートが完了しました。", vbInformation
    Exit Sub

Err_Handler:
    If isInTrans Then wsTrans.Rollback
    MsgBox "インポートプロセスでエラーが発生しました：" & vbCrLf & Err.Description, vbCritical
    
    If Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.Quit
        Set xlApp = Nothing
    End If
End Sub

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Pick_File
' 概要           : ファイル選択ダイアログを表示し、ユーザーにインポート対象ファイルを選択させます。
' 戻り値         : 選択されたファイルのフルパス（キャンセル時は空文字）
'---------------------------------------------------------------------------------------------------
Private Function Pick_File() As String
    Dim fd As Object: Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
        .title = "工事決定累積データ(Excel)を選択してください"
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx;*.xlsm"
        .AllowMultiSelect = False
        If .Show = -1 Then Pick_File = .SelectedItems(1)
    End With
End Function
