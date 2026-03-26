Attribute VB_Name = "acc_mod_Genka_Main"
'-------------------------------------
' Module: acc_mod_Genka_Main
' 説明   : 工事原価データのインポート、転送、枝番更新を一括管理する
'-------------------------------------
Option Compare Database
Option Explicit

'----------------------------------------------------------------
' サブルーチン名 : Import_GenkaData_ToMain
' 概要           : 仮テーブル at_Genka_Temp から本テーブルへデータを振り分けて転送する
'----------------------------------------------------------------
Public Sub Import_GenkaData_ToMain()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim dictDefault As Object
    Dim transfer As New acc_clsTableTransfer
    
    ' 1. デフォルトルール辞書の取得（at_Genka_DefaultRulesテーブルより）
    Set dictDefault = Get_DefaultRuleDictionary("at_Genka_DefaultRules")
    
    ' --- A. 基本工事の転送 ---
    ' 本テーブルをクリアしてから、行分類が「基本工事」のものを抽出して転送
    db.Execute "DELETE * FROM at_Genka_Kihon", dbFailOnError
    transfer.Init "at_Genka_Temp", "at_Genka_Kihon"
    transfer.Criteria = "[行分類] = '基本工事'"
    Set transfer.DefaultRules = dictDefault
    transfer.ExecuteTransfer
    
    ' --- B. 枝番工事の転送 ---
    ' 本テーブルをクリアしてから、枝番コードがあり「決定」済みのものを抽出して転送
    db.Execute "DELETE * FROM at_Genka_Edaban", dbFailOnError
    transfer.Init "at_Genka_Temp", "at_Genka_Edaban"
    transfer.Criteria = "[枝番工事コード] Is Not Null AND [状況] = '決定'"
    Set transfer.DefaultRules = dictDefault
    transfer.ExecuteTransfer
    
    ' 概要: at_Genka_Edaban_Fixテーブルに基づき、枝番コードを強制補正する
    Call Apply_EdabanCode_Fix
    
    MsgBox "工事原価データの転記が完了しました。", vbInformation
End Sub ' ← Import_GenkaData_ToMain 終了

'----------------------------------------------------------------
' サブルーチン名 : Apply_EdabanCode_Fix
' 概要           : at_Genka_Edaban_Fixテーブルに基づき、枝番コードを強制補正する
'----------------------------------------------------------------
Public Sub Apply_EdabanCode_Fix()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim strSQL As String
    
    ' 工事コードと管理番号が一致するレコードの枝番を上書き
    strSQL = "UPDATE at_Genka_Edaban " & _
             "INNER JOIN at_Genka_Edaban_Fix " & _
             "ON (at_Genka_Edaban.工事コード = at_Genka_Edaban_Fix.工事コード) " & _
             "AND (at_Genka_Edaban.管理番号 = at_Genka_Edaban_Fix.管理番号) " & _
             "SET at_Genka_Edaban.枝番工事コード = [at_Genka_Edaban_Fix].[枝番コード];"
    
    On Error GoTo ErrHandler
    db.Execute strSQL, dbFailOnError
    
    Debug.Print "枝番コードの強制補正が完了しました。"
    Exit Sub

ErrHandler:
    MsgBox "補正処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub


'----------------------------------------------------------------
' サブルーチン名 : Update_EdabanCode_ByReference
' 概要           : Icube累計（at_Icube_Archive）を参照し、枝番コードを自動更新する
'----------------------------------------------------------------
Public Sub Update_EdabanCode_ByReference()
    Const NAME_TAIL_LEN As Long = 10 ' 名称末尾の比較文字数
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsTarget As DAO.Recordset
    Dim rsRef As DAO.Recordset
    Dim dictRef As Object: Set dictRef = CreateObject("Scripting.Dictionary")
    Dim key As String

    ' 1. 参照辞書（Icube累計）の作成
    Set rsRef = db.OpenRecordset("at_Icube_Archive", dbOpenSnapshot)
    Do Until rsRef.EOF
        ' キー作成：工事コード | 通貨形式価格 | 工事名称末尾10文字
        key = Nz(rsRef!工事コード, "") & "|" & _
              FormatCurrency(Nz(rsRef!工事価格, 0)) & "|" & _
              Right(Nz(rsRef!追加工事名称, ""), NAME_TAIL_LEN)
        
        dictRef(key) = Nz(rsRef!枝番工事コード, "")
        rsRef.MoveNext
    Loop
    rsRef.Close

    ' 2. 枝番テーブルの更新
    Set rsTarget = db.OpenRecordset("at_Genka_Edaban", dbOpenDynaset)
    Do Until rsTarget.EOF
        key = Nz(rsTarget!工事コード, "") & "|" & _
              FormatCurrency(Nz(rsTarget!工事価格, 0)) & "|" & _
              Right(Nz(rsTarget!追加工事名称, ""), NAME_TAIL_LEN)
        
        If dictRef.Exists(key) Then
            rsTarget.Edit
            rsTarget!枝番工事コード = dictRef(key)
            rsTarget.Update ' ← 変更を確定させるにゃ
        End If
        rsTarget.MoveNext
    Loop
    rsTarget.Close

    MsgBox "枝番工事コードの自動更新が完了しました。", vbInformation
End Sub ' ← Update_EdabanCode_ByReference 終了

'----------------------------------------------------------------
' サブルーチン名 : Check_EdabanCode_Error
' 概要           : Icube累計と整合しない枝番コードに警告を記録する
'----------------------------------------------------------------
Public Sub Check_EdabanCode_Error()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsTarget As DAO.Recordset
    Dim rsRef As DAO.Recordset
    Dim dictRef As Object: Set dictRef = CreateObject("Scripting.Dictionary")
    Dim key As String

    ' 1. エラーフィールドを初期化（実際のフィールド名に合わせてください）
    db.Execute "UPDATE at_Genka_Edaban SET [枝番工事コードエラーチェック] = Null", dbFailOnError

    ' 2. 参照辞書の作成
    Set rsRef = db.OpenRecordset("at_Icube_Archive", dbOpenSnapshot)
    Do Until rsRef.EOF
        key = Nz(rsRef!枝番工事コード, "") & "|" & Nz(rsRef!追加工事名称, "")
        dictRef(key) = True
        rsRef.MoveNext
    Loop
    rsRef.Close

    ' 3. 不一致チェックの実行
    Set rsTarget = db.OpenRecordset("at_Genka_Edaban", dbOpenDynaset)
    Do Until rsTarget.EOF
        key = Nz(rsTarget!枝番工事コード, "") & "|" & Nz(rsTarget!追加工事名称, "")
        If Not dictRef.Exists(key) Then
            rsTarget.Edit
            rsTarget![枝番工事コードエラーチェック] = "枝番コード不一致の疑い"
            rsTarget.Update
        End If
        rsTarget.MoveNext
    Loop
    rsTarget.Close

    MsgBox "エラーチェックが完了しました。", vbInformation
End Sub ' ← Check_EdabanCode_Error 終了

'----------------------------------------------------------------
' 関数名 : Get_DefaultRuleDictionary
' 概要   : ルール管理テーブルから変換・補完ルールを読み込んで返す
'----------------------------------------------------------------
Private Function Get_DefaultRuleDictionary(ByVal TableName As String) As Object
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    Set rs = db.OpenRecordset(TableName, dbOpenSnapshot)
    Do Until rs.EOF
        ' 取込フラグが True の項目のみ対象とする
        If rs!取込フラグ = True Then
            ' dict(フィールド名) = Array(データ型, 空欄対応モード)
            dict(rs!accテーブル名) = Array(rs!データ型, rs!空欄対応モード)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set Get_DefaultRuleDictionary = dict
End Function ' ← Get_DefaultRuleDictionary 終了

