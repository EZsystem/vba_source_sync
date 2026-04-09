Attribute VB_Name = "acc_mod_Kenmu_Export"
'-------------------------------------
' Module: acc_mod_Kenmu_Export
' 概要  : 集計結果を元ファイルのExcelテーブルへ書き戻す
'-------------------------------------
Option Explicit

Private Const SRC_QUERY As String = "sqlSum_兼務率職員単位月毎"
Private Const TGT_SHEET As String = "集計"
Private Const TGT_LISTOBJ As String = "xt_兼務率職員単位月毎"

'--------------------------------------------
' プロシージャ名： Export_Sum_To_Excel
' 概要： クエリの結果を元ファイルパスごとに分類してExcelへ書き出す
'--------------------------------------------
Public Sub Export_Sum_To_Excel(Optional ByVal callingID As Long = 0)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsFiles As DAO.Recordset
    Dim rsData As DAO.Recordset
    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim lo As Object
    Dim filePath As String
    Dim sql As String
    
    ' 1.5 レジストリから出力先フォルダを取得
    Dim exportRoot As String: exportRoot = ""
    If callingID > 0 Then
        Set rsFiles = db.OpenRecordset("SELECT [既定パス] FROM [" & AT_SYSTEM_REG & "] WHERE ID = " & callingID, dbOpenSnapshot)
        If Not rsFiles.EOF Then
            exportRoot = Nz(rsFiles![既定パス], "")
        End If
        rsFiles.Close
    End If

    On Error GoTo ErrLine
    
    ' 2. ユニークな元ファイルパスのリストを取得
    Set rsFiles = db.OpenRecordset( _
        "SELECT 元ファイルパス FROM [" & SRC_QUERY & "] GROUP BY 元ファイルパス", dbOpenSnapshot)
    
    If rsFiles.EOF Then
        MsgBox "出力対象のデータが見つかりません。", vbInformation
        GoTo CleanUp
    End If
    
    ' 3. ファイルごとのループ処理
    Do Until rsFiles.EOF
        filePath = rsFiles!元ファイルパス
        
        ' --- 既定パスによる出力先のリダイレクト ---
        If exportRoot <> "" Then
            Dim fileNameOnly As String
            fileNameOnly = Get_FileName_EZ(filePath)
            ' 既定パスとファイル名を結合
            filePath = exportRoot & IIf(Right(exportRoot, 1) = "\", "", "\") & fileNameOnly
        End If
        
        ' ファイルが存在するか確認
        If Dir(filePath) <> "" Then
            Set wb = xlApp.Workbooks.Open(filePath)
            Set ws = Nothing
            On Error Resume Next
            Set ws = wb.Worksheets(TGT_SHEET)
            On Error GoTo ErrLine
            
            If Not ws Is Nothing Then
                Set lo = ws.ListObjects(TGT_LISTOBJ)
                If Not lo Is Nothing Then
                    
                    ' --- A. テーブルのクリア（1行残し） ---
                    ' DataBodyRangeがある場合のみ処理
                    If Not lo.DataBodyRange Is Nothing Then
                        If lo.ListRows.count > 1 Then
                            ' 2行目以降を削除
                            lo.DataBodyRange.Offset(1, 0).Resize(lo.ListRows.count - 1).Delete
                        End If
                        ' 1行目の内容をクリア（書式や数式維持のためDeleteではなくClearContents）
                        lo.DataBodyRange.rows(1).ClearContents
                    End If
                    
                    ' --- B. 対象ファイル用のデータを抽出 ---
                    ' パスでフィルタリングしたレコードセットを開く
                    sql = "SELECT 作業所名, 工事コード, 工事名, 社員名, 想定給与額, 総合職数, 人員数 " & _
                          "FROM [" & SRC_QUERY & "] WHERE 元ファイルパス = '" & Replace(filePath, "'", "''") & "'"
                    Set rsData = db.OpenRecordset(sql, dbOpenSnapshot)
                    
                    ' --- C. データの貼り付け ---
                    If Not rsData.EOF Then
                        ' テーブルの1行目（A1セル相当）から貼り付け
                        lo.InsertRowRange.Range("A1").CopyFromRecordset rsData
                    End If
                    rsData.Close
                    
                Else
                    Debug.Print "テーブルが見つかりません: " & filePath
                End If
            End If
            
            wb.Close SaveChanges:=True
        Else
            Debug.Print "ファイルが見つかりません: " & filePath
        End If
        
        rsFiles.MoveNext
    Loop
    
    Call Notify_Smart_Popup("全てのファイルへの書き出しが完了しました。", "完了通知")

CleanUp:
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    Call Fast_Mode_Toggle(False)
    Exit Sub

ErrLine:
    MsgBox "エラー発生: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
