Attribute VB_Name = "acc_mod_Genka_Main"
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
    Dim dictRules   As Object
    Dim db          As DAO.Database: Set db = CurrentDb
    
    Set clsLog = New com_clsErrorUtility
    clsLog.Init isBatch:=True
    
    Set clsTransfer = New acc_clsTableTransfer
    clsTransfer.Init db
    
    On Error GoTo Err_Handler
    
    ' 1. マッピングルールの取得
    Set dictRules = Get_GenkaRuleDictionary(AT_GENKA_COL_SETTING)
    
    ' --- A. 基本工事の転送 ---
    db.Execute "DELETE * FROM [" & AT_GENKA_BASIC & "]", dbFailOnError
    clsTransfer.ExecuteTransferWithRules AT_GENKA_IMPORT_WORK, AT_GENKA_BASIC, dictRules, "[行分類] = '基本工事'"
    
    ' --- B. 枝番工事の転送 ---
    db.Execute "DELETE * FROM [" & AT_GENKA_BRANCH & "]", dbFailOnError
    clsTransfer.ExecuteTransferWithRules AT_GENKA_IMPORT_WORK, AT_GENKA_BRANCH, dictRules, "[枝番工事コード] Is Not Null AND [状況] = '決定'"
    
    ' --- C. 手動最終補正の実行 (★ここが修正の核) ---
    ' 枝番工事コードを軸に、工事コード・管理番号・追加工事名称を正しい組み合わせに入れ替える
    Call Apply_Manual_Final_Correction(clsLog)
    
    clsLog.Notify_Smart_Popup "工事原価データの転記および最終補正が完了しました。"
    Exit Sub

Err_Handler:
    clsLog.Notify_Smart_Popup "Import_GenkaData Error: " & Err.Description
End Sub


' --- ルール辞書の生成 (列名を現状のテーブル構造に合わせる) ---
Private Function Get_GenkaRuleDictionary(ByVal tblName As String) As Object
    Dim rs As DAO.Recordset
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    ' 列名を修正: [取込フラグ], [accテーブル名], [データ型], [空欄対応モード]
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [" & tblName & "] WHERE [取込フラグ] = True", dbOpenSnapshot)
    Do Until rs.EOF
        ' dict(フィールド名) = Array(型名, デフォルト値/空欄対応モード)
        ' キーには [accテーブル名] を使用
        dict(rs![accテーブル名].Value) = Array(rs![データ型].Value, rs![空欄対応モード].Value)
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

