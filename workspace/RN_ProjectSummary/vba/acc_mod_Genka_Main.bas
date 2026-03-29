Attribute VB_Name = "acc_mod_Genka_Main"
'Attribute VB_Name = "acc_mod_Genka_Main"
'----------------------------------------------------------------
' サブルーチン名 : Import_GenkaData_ToMain
' 概要 : 原価データのインポート・転送、および各種補正（自動・手動）を一括実行する
'----------------------------------------------------------------
Option Compare Database
Option Explicit

'----------------------------------------------------------------
' サブルーチン名 : Import_GenkaData_ToMain
'----------------------------------------------------------------
Public Sub Import_GenkaData_ToMain()
    Dim clsLog      As com_clsErrorUtility
    Dim clsTransfer As acc_clsTableTransfer
    Dim dictBasic   As Object
    Dim dictBranch  As Object
    Dim db          As DAO.Database: Set db = CurrentDb
    
    Set clsLog = New com_clsErrorUtility
    clsLog.Init isBatch:=True
    
    Set clsTransfer = New acc_clsTableTransfer
    clsTransfer.Init db
    
    On Error GoTo Err_Handler
    
    Const TBL_SETTING As String = "at_原価S_ColSetting2"
    
    ' 1. マッピングルールの取得 (Test2/Test3で別々の辞書を作る)
    Set dictBasic = Get_GenkaRuleDictionary(TBL_SETTING, "基本工事_本テーブルタイトル名", "基本工事_データ型")
    Set dictBranch = Get_GenkaRuleDictionary(TBL_SETTING, "枝番工事_本テーブルタイトル名", "枝番工事_データ型")
    
    ' 2. テスト用テーブルの本番列名マッピングクエリ（仮想ビュー）を動的作成
    Const QRY_TEST2 As String = "qry_Genka_Map_Test2"
    Const QRY_TEST3 As String = "qry_Genka_Map_Test3"
    Call Create_Mapped_Query("at_Test2_Raw_Import", QRY_TEST2, TBL_SETTING, "基本工事_仮テーブルタイトル名", "基本工事_本テーブルタイトル名")
    Call Create_Mapped_Query("at_Test3_Raw_Import", QRY_TEST3, TBL_SETTING, "枝番工事_仮テーブルタイトル名", "枝番工事_本テーブルタイトル名")
    
    ' --- A. 基本工事の転送 ---
    db.Execute "DELETE * FROM [" & AT_GENKA_BASIC & "]", dbFailOnError
    ' Test2テーブルはすべて基本工事のデータのため、抽出条件は不要（行分類列もないため空文字を指定します）
    clsTransfer.ExecuteTransferWithRules QRY_TEST2, AT_GENKA_BASIC, dictBasic, ""
    
    ' --- B. 枝番工事の転送 ---
    db.Execute "DELETE * FROM [" & AT_GENKA_BRANCH & "]", dbFailOnError
    ' 枝番テーブルもTest3ですでに抽出済みのため、条件を外して全件取り込みます
    clsTransfer.ExecuteTransferWithRules QRY_TEST3, AT_GENKA_BRANCH, dictBranch, ""
    
    ' --- C. 手動最終補正の実行 (枝番工事コードを軸に属性差し替え) ---
    Call Apply_Manual_Final_Correction(clsLog)
    
    ' 後片付け (仮想クエリの削除)
    On Error Resume Next
    db.QueryDefs.Delete QRY_TEST2
    db.QueryDefs.Delete QRY_TEST3
    On Error GoTo Err_Handler
    
    clsLog.Notify_Smart_Popup "工事原価データの転記および最終補正が完了しました。"
    Exit Sub

Err_Handler:
    clsLog.Notify_Smart_Popup "Import_GenkaData Error: " & Err.Description
End Sub

'----------------------------------------------------------------
' サブルーチン名 : Create_Mapped_Query
' 概要 : 設定テーブル上の [仮テーブルタイトル名(F1等)] と [本テーブルタイトル名] のペアから仮想クエリを生成する
'----------------------------------------------------------------
Private Sub Create_Mapped_Query(ByVal sourceTable As String, ByVal queryName As String, ByVal tblSetting As String, ByVal colTempName As String, ByVal colMainName As String)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim strSelect As String
    Dim strSQL As String
    
    ' 列設定情報を読み込み SELECT リストを生成
    Set rs = db.OpenRecordset("SELECT * FROM [" & tblSetting & "]", dbOpenSnapshot)
    Do Until rs.EOF
        ' 本テーブルタイトル名（転送先列名）に文字が入っていればマッピング対象
        If Not IsNull(rs(colTempName).Value) And Not IsNull(rs(colMainName).Value) Then
            If Trim(rs(colTempName).Value & "") <> "" And Trim(rs(colMainName).Value & "") <> "" Then
                ' 例: [F1] AS [基本工事コード]
                strSelect = strSelect & "[" & rs(colTempName).Value & "] AS [" & rs(colMainName).Value & "], "
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    If Len(strSelect) > 0 Then
        strSelect = Left(strSelect, Len(strSelect) - 2)
    Else
        Exit Sub
    End If
    
    strSQL = "SELECT " & strSelect & " FROM [" & sourceTable & "];"
    
    ' 既存のクエリがある場合は削除
    On Error Resume Next
    db.QueryDefs.Delete queryName
    On Error GoTo 0
    
    ' 仮想クエリ(View)の作成
    Set qdf = db.CreateQueryDef(queryName, strSQL)
End Sub


' --- ルール辞書の生成 (列名を現状のテーブル構造に合わせる) ---
Private Function Get_GenkaRuleDictionary(ByVal tblName As String, ByVal colMainName As String, ByVal colTypeName As String) As Object
    Dim rs As DAO.Recordset
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [" & tblName & "]", dbOpenSnapshot)
    Do Until rs.EOF
        If Not IsNull(rs(colMainName).Value) Then
            If Trim(rs(colMainName).Value & "") <> "" Then
                ' 空欄対応モードは削られたため、一律でNullを渡す
                dict(rs(colMainName).Value) = Array(rs(colTypeName).Value, Null)
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set Get_GenkaRuleDictionary = dict
End Function


'----------------------------------------------------------------
' サブルーチン名 : Apply_Manual_Final_Correction
' 概要 : 枝番工事コードを検索軸として、テレコになった属性情報を
'        レコード単位で完全に差し替える（スワップ・ロジック）
'----------------------------------------------------------------
Private Sub Apply_Manual_Final_Correction(ByRef clsLog As com_clsErrorUtility)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim strSQL As String
    Const TEMP_TABLE As String = "at_Temp_Genka_Correction_Work"

    On Error GoTo Err_Sub

    ' 1. 作業用テーブルのクリーンアップ（存在すれば削除）
    On Error Resume Next
    db.Execute "DROP TABLE [" & TEMP_TABLE & "]", dbFailOnError
    On Error GoTo Err_Sub

    ' 2. 【退避】枝番工事コードを軸に、該当レコードの「全属性（30項目以上）」を一時テーブルへ退避
    ' [枝番工事コード](本テーブル) と [枝番コード](マスタ) でマッチング
    strSQL = "SELECT B.* INTO [" & TEMP_TABLE & "] " & _
             "FROM [" & AT_GENKA_BRANCH & "] AS B " & _
             "INNER JOIN [" & AT_GENKA_MANUAL_FIX & "] AS M " & _
             "ON B.[枝番工事コード] = M.[枝番コード];"
    db.Execute strSQL, dbFailOnError

    ' 3. 【属性修正】一時テーブル内の「工事コード」「管理番号」「追加工事名称」をマスタの正解値で上書き
    ' これにより、名札(枝番)はそのままで、属性の組み合わせだけが正しい状態になる
    strSQL = "UPDATE [" & TEMP_TABLE & "] AS T " & _
             "INNER JOIN [" & AT_GENKA_MANUAL_FIX & "] AS M " & _
             "ON T.[枝番工事コード] = M.[枝番コード] " & _
             "SET T.[工事コード] = M.[工事コード], " & _
             "    T.[管理番号] = M.[管理番号], " & _
             "    T.[追加工事名称] = M.[追加工事名称];"
    db.Execute strSQL, dbFailOnError

    ' 4. 【本テーブルの削除】差し替え対象となる枝番のレコードを本テーブルから一旦削除
    strSQL = "DELETE FROM [" & AT_GENKA_BRANCH & "] " & _
             "WHERE [枝番工事コード] IN (SELECT [枝番コード] FROM [" & AT_GENKA_MANUAL_FIX & "]);"
    db.Execute strSQL, dbFailOnError

    ' 5. 【復元】属性が正しく整理されたレコードを、一時テーブルから本テーブルへ一括挿入
    ' SELECT * を使用するため、フィールドが30個あってもパラメータエラーは発生しない
    strSQL = "INSERT INTO [" & AT_GENKA_BRANCH & "] SELECT * FROM [" & TEMP_TABLE & "];"
    db.Execute strSQL, dbFailOnError

    ' 6. 後片付け
    db.Execute "DROP TABLE [" & TEMP_TABLE & "]", dbFailOnError
    
    Debug.Print "枝番工事コードを軸としたデータ属性の入れ替えが完了しました。"
    Exit Sub

Err_Sub:
    clsLog.Notify_Smart_Popup "Apply_Manual_Final_Correction Error: " & Err.Description
End Sub



