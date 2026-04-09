Attribute VB_Name = "acc_mod_Import_Main"
Option Explicit

'----------------------------------------------------------------
' Module: acc_mod_Import_Main (EZsystem Integrated)
' 説明   : 職員兼務率の「Excel取込」と「本番転写」を一気に実行する統合モジュール
' 修正内容:
'    1. 定数管理を acc_mod_MappingTemplate へ集約
'    2. インポート終了後に自動で整形転写処理を開始
'    3. 全工程を単一のトランザクションで制御
'----------------------------------------------------------------

'--------------------------------------------
' プロシージャ名： Run_Kenmu_Import_EZ
' 概要： Excelから暫定テーブルへ取込後、整形して本番テーブルへ転送する
'--------------------------------------------
Public Sub Run_Kenmu_Import_EZ(Optional ByVal callingID As Long = 0)
    Dim db            As DAO.Database: Set db = CurrentDb
    Dim rsConfig      As DAO.Recordset
    Dim importer      As New acc_clsExcelImporter
    Dim xlApp         As Object
    Dim wb            As Object
    Dim ws            As Object
    Dim fso           As Object
    Dim selectedFiles As Collection
    Dim fileItem      As Variant
    Dim fileName      As String
    Dim inputFolder   As String
    Dim fileCount     As Long
    
    On Error GoTo ErrLine
    
    ' 1. レジストリから初期パス取得
    Dim strSQL As String
    If callingID > 0 Then
        strSQL = "SELECT [既定パス] FROM [" & AT_SYSTEM_REG & "] WHERE [ID] = " & callingID
    Else
        strSQL = "SELECT [既定パス] FROM [" & AT_SYSTEM_REG & "] WHERE [処理名称] = '職員兼務率インポート'"
    End If
    
    Set rsConfig = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rsConfig.EOF Then
        MsgBox "システムレジストリに設定が見つかりません (ID: " & callingID & ")", vbCritical
        Exit Sub
    End If
    inputFolder = Nz(rsConfig![既定パス], ""): rsConfig.Close
    
    ' 2. ユーザーによる複数ファイル選択
    Set selectedFiles = SelectMultipleFiles(inputFolder)
    If selectedFiles.count = 0 Then Exit Sub ' キャンセル時
    
    ' ★追加：パスの学習（選択したファイルの親フォルダをレジストリへ自動保存）
    If callingID > 0 And selectedFiles.count > 0 Then
        Dim newFolder As String
        newFolder = Left(selectedFiles(1), InStrRev(selectedFiles(1), "\"))
        If newFolder <> inputFolder Then
            Call Sync_Registry_Path(callingID, newFolder)
        End If
    End If

    
    ' 3. 初期準備
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set xlApp = CreateObject("Excel.Application")
    Call Fast_Mode_Toggle(True, xlApp)
    
    ' --- トランザクション開始 ---
    Dim wsJK As DAO.Workspace: Set wsJK = DBEngine.Workspaces(0)
    wsJK.BeginTrans
    
    ' 第一工程：暫定テーブルのクリアとインポート
    db.Execute "DELETE * FROM [" & AT_KENMU_TEMP & "];", dbFailOnError
    
    importer.Init db
    importer.TempTableName = AT_KENMU_TEMP
    
    For Each fileItem In selectedFiles
        fileName = fso.GetFileName(fileItem)
        fileCount = fileCount + 1
        
        Set wb = xlApp.Workbooks.Open(fileItem, ReadOnly:=True)
        
        ' シート特定（シート名またはオブジェクト名 Sheet1 で検索）
        Set ws = G_GetSheetByCodeName(wb, SH_NAME_KENMU)
        
        If ws Is Nothing Then Err.Raise 999, , "シート「" & SH_NAME_KENMU & "」が見つかりません: " & fileName
        
        Dim lo As Object
        On Error Resume Next
        Set lo = ws.ListObjects(LO_NAME_KENMU)
        On Error GoTo ErrLine
        
        If lo Is Nothing Then Err.Raise 999, , "テーブル「" & LO_NAME_KENMU & "」が見つかりません: " & fileName
        
        ' 作業所名チェック
        Dim worksiteName As String
        worksiteName = Trim(Nz(ws.Range("D2").Value, ""))
        If worksiteName = "" Then
            Err.Raise 999, , "作業所名（D2セル）が未入力です: " & fileName
        End If
        
        Call importer.ImportUnpivotedData(lo, worksiteName, CStr(fileItem))
        
        wb.Close SaveChanges:=False
        Set ws = Nothing
    Next fileItem
    
    ' 第二工程：暫定テーブル -> 本番テーブルへの整形・転送
    Call Transcribe_Integrated_Logic(db)
    
    ' 第三工程：本番テーブル -> 累計履歴への同期（転送元勝ち）
    Call Update_Kenmu_History(db)
    
    ' --- すべて成功したら確定 ---
    wsJK.CommitTrans
    
    Call Notify_Smart_Popup(fileCount & " 件のファイルをインポート・累計同期しました。", "完了通知")

CleanUp:
    On Error Resume Next
    Call Fast_Mode_Toggle(False, xlApp)
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    Exit Sub

ErrLine:
    ' エラー情報を即座に保存
    Dim errNum As Long:   errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    
    On Error Resume Next
    If Not wsJK Is Nothing Then wsJK.Rollback
    
    ' 特定のエラー番号に対する詳細メッセージ
    Dim customMsg As String
    If errNum = 3022 Then
        customMsg = "【ID重複エラー】既に取り込まれている累計データと ID が重複しました。一度テーブルをクリアするか、時間を置いて再試行してください。"
    ElseIf errNum = 3078 Then
        customMsg = "【テーブル不在】インポート先のテーブル名が正しくありません。"
    Else
        customMsg = "エラー内容(" & errNum & "): " & errDesc
    End If
    
    MsgBox "【インポート中断】" & vbCrLf & _
           "ファイル: " & fileName & vbCrLf & _
           "処理位置: " & IIf(ws Is Nothing, "シート確認中", "データ読込中") & vbCrLf & _
           customMsg, vbCritical
    Resume CleanUp
End Sub

'--------------------------------------------
' 統合ロジック：暫定から本番への整形・転回
'--------------------------------------------
Private Sub Transcribe_Integrated_Logic(ByRef db As DAO.Database)
    Dim rsSrc As DAO.Recordset
    Dim rsTgt As DAO.Recordset
    Dim arrMap As Variant
    Dim objOrg As Object
    Dim m      As Long
    
    ' 仮基本工事マッピングおよび組織マッピングを事前にロード
    arrMap = Get_TempProject_Map()
    Set objOrg = Get_Org_Dict()
    
    ' 本番テーブルをクリア
    db.Execute "DELETE * FROM [" & AT_KENMU_MAIN & "];", dbFailOnError
    
    Set rsSrc = db.OpenRecordset(AT_KENMU_TEMP, dbOpenSnapshot)
    Set rsTgt = db.OpenRecordset(AT_KENMU_MAIN, dbOpenDynaset)
    
    Do Until rsSrc.EOF
        ' 兼務率の正規化
        Dim dblRate As Double
        dblRate = Cleanse_Percent_Smart(rsSrc!兼務率割合)
        
        ' 値が 0（またはエラー）でない場合のみ本番へ登録
        If dblRate <> 0 Then
            rsTgt.AddNew
            
            ' 主キー（ImportID）の引き継ぎ
            rsTgt!ImportID = rsSrc!ImportID
            
            rsTgt!元ファイルパス = rsSrc!元ファイルパス
            
            ' 作業所名の正規化（略称から正式組織名に置換）
            Dim sWorksite As String: sWorksite = Nz(rsSrc!作業所名, "")
            Dim sOrgKey   As String: sOrgKey = StrConv(sWorksite, vbNarrow)
            
            If objOrg.Exists(sOrgKey) Then
                rsTgt!作業所名 = objOrg(sOrgKey)
            Else
                rsTgt!作業所名 = sWorksite
            End If
            
            rsTgt!No = rsSrc!No
            rsTgt!工事名 = rsSrc!工事名
            rsTgt!コメント = rsSrc!コメント
            rsTgt!社員名 = rsSrc!社員名
            
            ' 日付整形（MappingTemplateのロジックを使用）
            Dim dtFinal As Variant
            dtFinal = Cleanse_Date_Smart(rsSrc!年月)
            rsTgt!年月 = dtFinal
            rsTgt!期 = Get_FiscalTerm(dtFinal)
            rsTgt!Q = Get_Quarter(dtFinal)
            
            ' 仮基本工事コードの紐付け（ワイルドカード考慮）
            Dim projectName As String
            Dim tempCode    As String: tempCode = ""
            Dim sourceCode  As String: sourceCode = Trim(Nz(rsSrc!工事コード, ""))
            
            projectName = Nz(rsSrc!工事名, "")
            
            If IsArray(arrMap) Then
                For m = 0 To UBound(arrMap, 1)
                    ' arrMap(m, 0) には既に "？" -> "?" 変換済みのパターンが入っている
                    If projectName Like arrMap(m, 0) Then
                        tempCode = arrMap(m, 1)
                        ' ★追加：2文字目を四半期の数字(1-4)に書き換える (例: D1M -> D4M)
                        If Len(tempCode) >= 2 Then
                            Mid(tempCode, 2, 1) = Left(Nz(rsTgt!Q, "1"), 1)
                        End If
                        Exit For
                    End If
                Next m
            End If
            
            ' --- 【修正】工事コードのセット（未定補完あり） ---
            If (sourceCode = "未定" Or sourceCode = "") And tempCode <> "" Then
                rsTgt!工事コード = tempCode
            Else
                rsTgt!工事コード = sourceCode
            End If
            
            ' --- 【修正】仮基本工事コードのセット ---
            rsTgt!仮基本工事コード = IIf(tempCode <> "", tempCode, rsTgt!工事コード)
            
            rsTgt!兼務率割合 = dblRate
            
            rsTgt.Update
        End If
        rsSrc.MoveNext
    Loop
    
    rsSrc.Close
    rsTgt.Close
End Sub

'--------------------------------------------
' 累計同期ロジック：転送元勝ち（Delete & Insert）
'--------------------------------------------
Private Sub Update_Kenmu_History(ByRef db As DAO.Database)
    Dim strSQL As String
    
    ' 1. 同一キー（年月+工事コード+社員名+兼務率割合）を持つ既存レコードを削除
    strSQL = "DELETE FROM [" & AT_KENMU_HISTORY & "] " & _
             "WHERE EXISTS (" & _
             "  SELECT 1 FROM [" & AT_KENMU_MAIN & "] AS SRC " & _
             "  WHERE [" & AT_KENMU_HISTORY & "].[年月] = SRC.[年月] " & _
             "    AND [" & AT_KENMU_HISTORY & "].[工事コード] = SRC.[工事コード] " & _
             "    AND [" & AT_KENMU_HISTORY & "].[社員名] = SRC.[社員名] " & _
             "    AND [" & AT_KENMU_HISTORY & "].[兼務率割合] = SRC.[兼務率割合] " & _
             ")"
    db.Execute strSQL, dbFailOnError
    
    ' 2. 本番テーブルから累計テーブルへ全件追加
    ' ※ImportID が重複した場合は、上位の ErrLine でトラップされる
    strSQL = "INSERT INTO [" & AT_KENMU_HISTORY & "] SELECT * FROM [" & AT_KENMU_MAIN & "];"
    db.Execute strSQL, dbFailOnError
End Sub
