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
'   - テーブル   : tbl_xl_IcubeColSetting (列マッピング定義マスタ)
' 最終更新日     : 2026/03/26
'===================================================================================================

'---------------------------------------------------------------------------------------------------
' プロシージャ名 : Run_IcubeImport_FullStep
' 概要           : 外部Excelファイルを読み込み、定義に基づいて本番テーブル(Icube_)へデータを取り込みます。
' 処理フロー     : 1. Excel読込(物理置換) -> 2. 仮保存テーブル(at_Temp) -> 3. 本番転写(型変換)
' 引数           : なし
' 戻り値         : なし
'---------------------------------------------------------------------------------------------------
Public Sub Run_IcubeImport_FullStep()
    ' --- 1. 設定値定義（マジックナンバーの排除） ---
    ' ※EZsystem規約に基づき、テーブル参照時は at_ プレフィックスを推奨（現状維持）
    Const TBL_TEMP      As String = "tbl_Temp_Icube_Import"
    Const TBL_TARGET    As String = "Icube_"
    Const EXCEL_HEADER  As Long = 5      ' ヘッダー行番号
    Const EXCEL_START_C As Long = 2      ' 取込開始列（B列）
    Const EXCEL_END_C   As Long = 141    ' 取込終了列（EK列）
    Const KEY_FIELD     As String = "No" ' 重複排除および更新時のキーフィールド

    Dim Db          As DAO.Database: Set Db = CurrentDb
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
    If filePath = "" Then Exit Sub ' キャンセル時は中断

    ' --- 3. クラスの初期化とインポート設定 ---
    pImporter.Init Db
    pImporter.TableName = TBL_TEMP
    pImporter.HeaderRow = EXCEL_HEADER
    pImporter.StartColumn = EXCEL_START_C
    pImporter.EndColumn = EXCEL_END_C
    
    pTransfer.Init Db

    ' --- 4. マッピングルールのロード ---
    ' tbl_xl_IcubeColSettingから取込対象フィールドと置換後の名称を取得
    Set RsMaster = Db.OpenRecordset("SELECT [タイトル名_デフォルト], [タイトル名_置換え後] FROM tbl_xl_IcubeColSetting WHERE Nz(取込フラグ, False) = True")
    Do Until RsMaster.EOF
        ' Excelのヘッダー名とAccessのフィールド名の紐付けをクラスへ登録
        pImporter.AddMapping CStr(RsMaster![タイトル名_デフォルト]), CStr(RsMaster![タイトル名_置換え後])
        RsMaster.MoveNext
    Loop
    RsMaster.Close

    ' --- 5. データ整合性確保のためのトランザクション開始 ---
    wsTrans.BeginTrans
    isInTrans = True

    ' --- 6. 仮テーブルのクリア ---
    Db.Execute "DELETE FROM [" & TBL_TEMP & "];", dbFailOnError
    
    ' --- 7. Excelアプリケーション制御 ---
    Set xlApp = CreateObject("Excel.Application")
    Set wb = xlApp.Workbooks.Open(filePath, ReadOnly:=True)
    
    ' インポート実行（acc_clsExcelImporter内部で物理置換ロジックが動作）
    pImporter.ImportFromWorksheet wb.Sheets(1)
    
    ' Excelリソースの解放
    wb.Close False
    xlApp.Quit
    Set xlApp = Nothing

    ' --- 8. 本番テーブル(Icube_)への型変換転写 ---
    ' acc_clsTableTransferにより、空文字をNullとして扱い、既存データとの重複をKEY_FIELDで制御
    pTransfer.ExecuteTransfer TBL_TEMP, TBL_TARGET, KEY_FIELD

    ' --- 9. 処理の確定 ---
    wsTrans.CommitTrans
    isInTrans = False

    MsgBox "クラスによる統合インポートが完了しました。", vbInformation
    Exit Sub

Err_Handler:
    ' ロールバック処理：エラー発生時にデータの不整合を防ぐ
    If isInTrans Then wsTrans.Rollback
    MsgBox "インポートプロセスでエラーが発生しました：" & vbCrLf & Err.Description, vbCritical
    
    ' オブジェクトの安全な破棄
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
