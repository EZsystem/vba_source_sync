Attribute VB_Name = "mod_IcubeImport_Refined"
Option Compare Database
Option Explicit

'==========================================================
' 処理名 : Run_IcubeImport_FullStep
' 説明   : クラスを活用した統合インポートフロー
'          1. Excel読込(物理置換) -> 2. 仮保存 -> 3. 本番転写(型変換)
'==========================================================
Public Sub Run_IcubeImport_FullStep()
    ' --- 設定値（実数はここに集約） ---
    Const TBL_TEMP      As String = "tbl_Temp_Icube_Import"
    Const TBL_TARGET    As String = "Icube_"
    Const EXCEL_HEADER  As Long = 5
    Const EXCEL_START_C As Long = 2   ' B列
    Const EXCEL_END_C   As Long = 141 ' EK列
    Const KEY_FIELD     As String = "No" ' 重複削除用キー（昨日の成功に基づきNoを指定）

    Dim Db          As DAO.Database: Set Db = CurrentDb
    Dim wsTrans     As DAO.Workspace: Set wsTrans = DBEngine.Workspaces(0)
    
    ' クラスのインスタンス化
    Dim pImporter   As New acc_clsExcelImporter
    Dim pTransfer   As New acc_clsTableTransfer
    
    Dim xlApp       As Object
    Dim wb          As Object
    Dim filePath    As String
    Dim RsMaster    As DAO.Recordset
    Dim isInTrans   As Boolean: isInTrans = False
    
    On Error GoTo Err_Handler

    ' --- 1. ファイル選択 ---
    filePath = Pick_File()
    If filePath = "" Then Exit Sub

    ' --- 2. クラスの初期化と設定 ---
    pImporter.Init Db
    pImporter.TableName = TBL_TEMP
    pImporter.HeaderRow = EXCEL_HEADER
    pImporter.StartColumn = EXCEL_START_C
    pImporter.EndColumn = EXCEL_END_C
    
    pTransfer.Init Db

    ' --- 3. マスタから翻訳ルールをクラスに教える ---
    Set RsMaster = Db.OpenRecordset("SELECT [タイトル名_デフォルト], [タイトル名_置換え後] FROM tbl_xl_IcubeColSetting WHERE Nz(取込フラグ, False) = True")
    Do Until RsMaster.EOF
        pImporter.AddMapping CStr(RsMaster![タイトル名_デフォルト]), CStr(RsMaster![タイトル名_置換え後])
        RsMaster.MoveNext
    Loop
    RsMaster.Close

    ' ==========================================================
    ' トランザクション開始
    ' ==========================================================
    wsTrans.BeginTrans
    isInTrans = True

    ' --- 4. 仮テーブルクリア & Excel取込 ---
    Db.Execute "DELETE FROM [" & TBL_TEMP & "];", dbFailOnError
    
    Set xlApp = CreateObject("Excel.Application")
    Set wb = xlApp.Workbooks.Open(filePath, ReadOnly:=True)
    
    ' クラスにシートを渡して取込実行（物理置換ロジックが中で動く）
    pImporter.ImportFromWorksheet wb.Sheets(1)
    
    wb.Close False: xlApp.Quit: Set xlApp = Nothing

    ' --- 5. 本テーブル(Icube_)へ型変換を考慮して転写 ---
    ' 内部で IIf([列名]='', Null, [列名]) が走り、型エラーを回避します
    pTransfer.ExecuteTransfer TBL_TEMP, TBL_TARGET, KEY_FIELD

    ' --- 6. 確定 ---
    wsTrans.CommitTrans
    isInTrans = False

    MsgBox "クラスによる統合インポートが完了しましたにゃ！", vbInformation
    Exit Sub

Err_Handler:
    ' 失敗した場合は全て取り消し
    If isInTrans Then wsTrans.Rollback
    MsgBox "失敗しました： " & Err.Description, vbCritical
    If Not xlApp Is Nothing Then xlApp.Quit
End Sub

' --- ファイル選択ヘルパー ---
Private Function Pick_File() As String
    Dim fd As Object: Set fd = Application.FileDialog(3)
    With fd
        .title = "工事決定累積データ(Excel)を選択してください"
        .Filters.Clear: .Filters.Add "Excelファイル", "*.xlsx;*.xlsm"
        If .Show = -1 Then Pick_File = .SelectedItems(1)
    End With
End Function

