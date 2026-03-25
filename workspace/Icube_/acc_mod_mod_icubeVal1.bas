Attribute VB_Name = "mod_icubeVal1"
Option Compare Database
Option Explicit

'テーブル：Icube_の記載処理
Public Sub mod_icube_All_1()
'工事名称から一件工事判定
    Call mod_icube_input1
'基本コードがブランクの時、工事コードを転写
    Call mod_icube_copy1
'基本工事名称が無い場合に工事帳票名を転写
    Call mod_icube_copy2
'枝番工事コードの記入(工事コードと枝番の連結)
    Call mod_icube_merge1
'受注計上年月を使い、受注年月日等を記入
    Call mod_icube_dateCnv_1
'完工日枝番を使い、完工年月日等を記入
    Call mod_icube_dateCnv_2
'テーブルIcubeへ施工管轄組織コードの記入
    Call mod_icube_copy4
'テーブルIcubeへ施工管轄組織名の記入
    Call mod_icube_copy5
'テーブルIcubeへ基本工事名称を分割して記入
    Call mod_icube_Val2ALL


' 処理完了メッセージ
    'MsgBox "処理が完了しました。", vbInformation
End Sub



'============================================
' プロシージャ名 : mod_icube_copy4
' 概要           : テーブル「Icube_」に、施工管轄組織コードを記入する
'                  ・t_支店作業所_累計 の対応辞書で Icube_ の施工担当組織コードを置換
'                  ・該当しない場合は t_err作業所 へエラー記録
'============================================
Public Sub mod_icube_copy4()

    ' --- 変数・オブジェクト定義 ---
    Dim db As DAO.Database                     ' DB接続オブジェクト
    Dim rsTarget As DAO.Recordset              ' 対象: Icube_ テーブル
    Dim rsError As DAO.Recordset               ' エラー記録: t_err作業所
    Dim rsCheck As DAO.Recordset               ' 辞書生成用: t_支店作業所_累計
    Dim strSQLTarget As String                 ' Icube_用SQL
    Dim dict As Object                         ' 組織コード→管轄コードの辞書
    Dim key As String
    Dim isErrorExists As Boolean

    ' --- データベースの初期化 ---
    Set db = CurrentDb()

    ' --- 施工管轄組織コードの辞書を作成 ---
    Set dict = CreateObject("Scripting.Dictionary")
    Set rsCheck = db.OpenRecordset("SELECT * FROM t_支店作業所_累計", dbOpenSnapshot)

    Do While Not rsCheck.EOF
        key = Trim(CStr(rsCheck!組織コード))  ' 組織コード（キー）

        ' 未登録かつ施工管轄組織コードがNULLでない場合のみ登録
        If Not dict.Exists(key) Then
            If Not IsNull(rsCheck!施工管轄組織コード) Then
                dict.Add key, CStr(rsCheck!施工管轄組織コード)   ' 管轄コードを値に
            End If
        End If
        rsCheck.MoveNext
    Loop
    rsCheck.Close
    Set rsCheck = Nothing

    ' --- Icube_全件を取得 ---
    strSQLTarget = "SELECT * FROM Icube_;"
    Set rsTarget = db.OpenRecordset(strSQLTarget, dbOpenDynaset)

    ' --- エラー記録用テーブル取得（未登録作業所記録用） ---
    Set rsError = db.OpenRecordset("t_err作業所", dbOpenDynaset)

    isErrorExists = False   ' エラーフラグ初期化

    ' --- Icube_各レコードを走査 ---
    If Not rsTarget.EOF Then
        rsTarget.MoveFirst
        Do While Not rsTarget.EOF
            Dim orgCode As String
            Dim valueToWrite As Variant

            ' 施工担当組織コードを取得（Null対策）
            If Not IsNull(rsTarget!施工担当組織コード) Then
                orgCode = Trim(CStr(rsTarget!施工担当組織コード))
            Else
                orgCode = ""
            End If

            ' --- 組織コードが辞書に存在する場合 ---
            If dict.Exists(orgCode) Then
                valueToWrite = dict(orgCode)

                ' 管轄コードが空やNullでなければ Icube_ に反映
                If Not IsNull(valueToWrite) Then
                    If Len(Trim(valueToWrite & "")) > 0 Then
                        rsTarget.Edit
                        rsTarget!施工管轄組織コード = valueToWrite
                        rsTarget.Update
                    Else
                        Debug.Print "空文字だったにゃ：" & orgCode
                    End If
                Else
                    Debug.Print "Null値だったにゃ：" & orgCode
                End If

            ' --- 辞書に存在しない場合はエラー記録 ---
            Else
                rsError.AddNew
                rsError!追加工事名称 = rsTarget!追加工事名称
                rsError!施工担当組織コード = rsTarget!施工担当組織コード
                rsError!施工担当組織名 = rsTarget!施工担当組織名
                rsError.Update
                isErrorExists = True
                Debug.Print "辞書に見つからなかったにゃ：" & orgCode
            End If

            rsTarget.MoveNext
        Loop
    End If

    ' --- リソースクリーンアップ ---
    rsTarget.Close
    rsError.Close
    Set rsTarget = Nothing
    Set rsError = Nothing
    Set db = Nothing
    Set dict = Nothing

    ' --- エラーがあれば通知 ---
    If isErrorExists Then
        MsgBox "未登録作業所があります", vbExclamation, "エラー"
    End If

End Sub



'テーブル：Icubeのレコードクリア
Public Sub mod_icubeClear1()
    On Error GoTo ErrHandler

    ' データベースオブジェクトの宣言
    Dim db As DAO.Database
    Dim sql As String
    
    ' データベースを取得
    Set db = CurrentDb()
    
    ' レコードを全削除するSQL
    sql = "DELETE * FROM Icube_"
    
    ' クエリの実行
    db.Execute sql, dbFailOnError
    
    ' 処理成功のメッセージ
    MsgBox "テーブル『Icube_』のレコードを全てクリアしました！", vbInformation

ExitProcedure:
    ' 後処理
    On Error Resume Next
    Set db = Nothing
    Exit Sub

ErrHandler:
    ' エラー時の処理
    MsgBox "エラーが発生しました: " & Err.description, vbCritical
    Resume ExitProcedure
End Sub



'テーブルIcubeへ施工管轄組織名の記入
Public Sub mod_icube_copy5()

    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim strSQLSource As String
    Dim strSQLTarget As String
    Dim dict As Object

    ' データベース
    Set db = CurrentDb()
    
    ' データ取得SQL
    strSQLSource = "SELECT 施工管轄組織コード, 施工管轄組織名 FROM tb_管轄作業所_RN部恒久作業所3;"
    strSQLTarget = "SELECT 施工管轄組織コード, 施工管轄組織名 FROM Icube_;"

    ' Dictionary 作成
    Set dict = CreateObject("Scripting.Dictionary")

    ' 参照元を開いて辞書に格納
    Set rsSource = db.OpenRecordset(strSQLSource, dbOpenSnapshot)
    Do While Not rsSource.EOF
        dict(Trim(CStr(rsSource!施工管轄組織コード))) = rsSource!施工管轄組織名
        rsSource.MoveNext
    Loop
    rsSource.Close
    Set rsSource = Nothing

    ' 参照先を開いて更新処理
    Set rsTarget = db.OpenRecordset(strSQLTarget, dbOpenDynaset)
    If Not rsTarget.EOF Then
        rsTarget.MoveFirst
        Do While Not rsTarget.EOF
            Dim key As String
            key = Trim(CStr(rsTarget!施工管轄組織コード))

            rsTarget.Edit
            If dict.Exists(key) Then
                rsTarget!施工管轄組織名 = dict(key)
                'Debug.Print "更新: " & key & " => " & dict(key)
            Else
                rsTarget!施工管轄組織名 = Null
                Debug.Print "一致なし: " & key
            End If
            rsTarget.Update

            rsTarget.MoveNext
        Loop
    End If

    rsTarget.Close
    Set rsTarget = Nothing
    Set db = Nothing
    Set dict = Nothing

    'MsgBox "Icube_ への転写処理が完了したにゃ", vbInformation

End Sub

'=================================================
' サブルーチン名 : mod_icube_dateCnv_1
' 説明   : Icube_ の年月テキストから各日付項目を更新するにゃ
'=================================================
Public Sub mod_icube_dateCnv_1()
    On Error GoTo EH

    ' --- 初期化 ---
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim dateMath As New com_clsDateMath
    Dim rawText As String

    ' --- データ取得 ---
    Set rs = db.OpenRecordset( _
        "SELECT No, [データ年月（受注計上年月）], 受注年度, 受注期, 受注Q, 受注月, 受注計上日_日付型 " & _
        "FROM Icube_", dbOpenDynaset)

    ' --- レコードが存在する場合のみ処理 ---
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            rawText = Nz(rs![データ年月（受注計上年月）], "")

            ' --- 日付文字列をセットすると自動で解析されるにゃ ---
            dateMath.RawValue = rawText

            If dateMath.IsValid Then
                rs.Edit
                rs!受注年度 = dateMath.GetFiscalYear
                rs!受注期 = dateMath.GetPeriod
                rs!受注Q = dateMath.GetQuarter
                rs!受注月 = dateMath.GetMonth
                rs!受注計上日_日付型 = dateMath.GetDateValue
                rs.Update
            Else
                Debug.Print "※無効な年月：" & rawText & " (No: " & rs!No & ")"
            End If

            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set dateMath = Nothing

    'MsgBox "Icube_ テーブルの日付変換が完了したにゃ！", vbInformation
    Exit Sub

' --- エラーハンドリング ---
EH:
    MsgBox "エラーが発生したにゃ：" & vbCrLf & Err.description, vbCritical
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set dateMath = Nothing
End Sub


'=================================================
' サブルーチン名 : mod_icube_dateCnv_2
' 説明   : 完成年月日（枝番単位）から完工年度・期・Q・月・日付型を再算出して更新する
'=================================================
Public Sub mod_icube_dateCnv_2()
    On Error GoTo Err_Handler

    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim dateMath As New com_clsDateMath
    Dim rawText As String

    Set rs = db.OpenRecordset( _
        "SELECT No, [完成年月日（枝番単位）], 完工年度, 完工期, 完工Q, 完工月, 完工日_日付型 " & _
        "FROM Icube_", dbOpenDynaset)

    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            rawText = Nz(rs![完成年月日（枝番単位）], "")
            dateMath.RawValue = rawText

            If dateMath.IsValid Then
                rs.Edit
                rs!完工年度 = dateMath.GetFiscalYear
                rs!完工期 = dateMath.GetPeriod
                rs!完工Q = dateMath.GetQuarter
                rs!完工月 = dateMath.GetMonth
                rs!完工日_日付型 = dateMath.GetDateValue
                rs.Update
            Else
                Debug.Print "※無効な完成年月日：" & rawText & " (No: " & rs!No & ")"
            End If

            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set dateMath = Nothing

    'MsgBox "完工日データの更新が完了したにゃ！", vbInformation
    Exit Sub

Err_Handler:
    MsgBox "エラーが発生したにゃ：" & vbCrLf & Err.description, vbExclamation
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set dateMath = Nothing
End Sub



'枝番工事コードの記入(工事コードと枝番の連結)
Public Sub mod_icube_merge1()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim TableName As String
    Dim primaryKeyField As String
    Dim field1 As String
    Dim field2 As String
    Dim targetField As String
    Dim combinedValue As String

    ' 処理対象情報
    TableName = "Icube_"
    primaryKeyField = "No"
    field1 = "工事コード"
    field2 = "工事枝番"
    targetField = "枝番工事コード"

    ' データベースとレコードセットを取得
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT * FROM " & TableName, dbOpenDynaset)

    ' レコードをループ
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            ' 連結フィールド1と連結フィールド2の値を取得
            Dim value1 As Variant
            Dim value2 As Variant

            value1 = rs.Fields(field1).Value
            value2 = rs.Fields(field2).Value

            ' Null値を処理する（空文字に置換）
            If IsNull(value1) Then value1 = ""
            If IsNull(value2) Then value2 = ""

            ' フィールド値を連結
            combinedValue = value1 & "-" & value2

            ' 連結結果を記入フィールドに更新
            rs.Edit
            rs.Fields(targetField).Value = combinedValue
            rs.Update

            rs.MoveNext
        Loop
    End If

    ' クリーンアップ
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    'MsgBox "フィールド値の連結処理が完了しました。", vbInformation
End Sub


'基本コードがブランクの時、工事コードを転写
Public Sub mod_icube_copy1()
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim transferCount As Long
    transferCount = 0 ' 転写件数を初期化

    ' データベースを取得
    Set db = CurrentDb()

    ' 対象レコードを取得
    strSQL = "SELECT No, 工事コード, 基本工事コード FROM Icube_"
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)

    ' レコードをループ
    Do While Not rs.EOF
        ' 転写先フィールドがブランクまたはNullの場合
        If IsNull(rs!基本工事コード) Or rs!基本工事コード = "N/A" Then
            rs.Edit
            rs!基本工事コード = rs!工事コード
            rs.Update
            transferCount = transferCount + 1 ' 転写件数をカウント
        End If
        rs.MoveNext
    Loop

    ' 結果を表示
    'MsgBox "転写処理が完了しました。転写件数: " & transferCount & " 件", vbInformation

CleanUp:
    ' 後始末
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました: " & Err.description, vbExclamation
    Resume CleanUp
End Sub

'基本工事名称が無い場合に工事帳票名を転写
Public Sub mod_icube_copy2()
    ' 定義
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    ' データベースの参照
    Set db = CurrentDb
    
    ' 対象レコードのSQLを定義
    strSQL = "SELECT No, 工事帳票名, 基本工事名称 " & _
             "FROM Icube_ " & _
             "WHERE [基本工事名称] IS NULL OR [基本工事名称] = '' OR [基本工事名称] = 'N/A';"
    
    ' レコードセットを開く
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    ' レコードが存在する場合
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            ' 転写処理
            If Not IsNull(rs!工事帳票名) And rs!工事帳票名 <> "" Then
                rs.Edit
                rs!基本工事名称 = rs!工事帳票名
                rs.Update
            End If
            rs.MoveNext
        Loop
    End If
    
    ' クリーンアップ
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' 処理完了メッセージ
    'MsgBox "工事帳票名の転写処理が完了しました。", vbInformation
End Sub



'基本工事名称から一件工事判定
Public Sub mod_icube_input1()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim targetTable As String
    Dim targetFieldCondition As String
    Dim targetFieldUpdate As String
    Dim conditionValues As Variant
    Dim updateValue1 As String
    Dim updateValue2 As String
    
    ' 処理対象情報
    targetTable = "Icube_"
    targetFieldCondition = "基本工事名称"
    targetFieldUpdate = "一件工事判定"
    
    ' 条件値
    conditionValues = Array("１２諸工事", "１３諸工事", "１Ｑ", "２Ｑ", "３Ｑ", "４Ｑ")
    
    ' 記入値
    updateValue1 = "小口工事"
    updateValue2 = "一件工事"
    
    ' データベースの参照
    Set db = CurrentDb()
    
    ' 対象テーブルのデータを取得
    strSQL = "SELECT No, " & targetFieldCondition & ", " & targetFieldUpdate & " FROM " & targetTable
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    ' レコードのループ処理
    If Not rs.EOF Then
        Do While Not rs.EOF
            Dim currentCondition As String
            Dim shouldUpdate As Boolean
            Dim i As Integer
            
            currentCondition = Nz(rs.Fields(targetFieldCondition).Value, "")
            shouldUpdate = False
            
            ' 条件値に該当するか確認
            For i = LBound(conditionValues) To UBound(conditionValues)
                If InStr(1, currentCondition, conditionValues(i), vbTextCompare) > 0 Then
                    shouldUpdate = True
                    Exit For
                End If
            Next i
            
            ' フィールド値を更新
            If shouldUpdate Then
                rs.Edit
                rs.Fields(targetFieldUpdate).Value = updateValue1
                rs.Update
            Else
                rs.Edit
                rs.Fields(targetFieldUpdate).Value = updateValue2
                rs.Update
            End If
            
            rs.MoveNext
        Loop
    End If
    
    ' クリーンアップ
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    'MsgBox "処理が完了しました。", vbInformation
End Sub

'エラー値置換え
 ' エラー値置換え対象テーブルの "エラー値置換え有無=TRUE"が実施対象
Public Sub mod_icube_err_1()
    Dim db As DAO.Database
    Dim rsTarget As DAO.Recordset
    Dim rsIcube As DAO.Recordset
    
    Dim clsCom As New cls_err
    Dim targetFieldName As String
    Dim fieldType As Integer
    Dim oldValue As Variant
    Dim newValue As Variant
    
    Set db = CurrentDb
    
    ' エラー値置換え対象テーブルの "エラー値置換え有無=TRUE" のフィールドを取得
    Set rsTarget = db.OpenRecordset( _
        "SELECT [フィールド名] " & _
        "FROM t_エラー値置換え対象 " & _
        "WHERE [エラー値置換え有無] = TRUE" _
    )
    
    If rsTarget.EOF Then
        MsgBox "置換え対象のフィールドがありません。", vbInformation
        GoTo CleanUp
    End If
    
    ' t_エラー値置換え対象 をレコード単位で走査
    Do While Not rsTarget.EOF
        
        targetFieldName = rsTarget!フィールド名
        
        ' Icube_ テーブルの該当フィールドを取得
        ' ※必要に応じて必要なフィールドをSELECT句で指定して下さい
        Set rsIcube = db.OpenRecordset( _
            "SELECT [No], [" & targetFieldName & "] " & _
            "FROM Icube_" _
        )
        
        ' Icube_の全レコードをループ
        Do While Not rsIcube.EOF
            
            oldValue = rsIcube.Fields(targetFieldName).Value
            fieldType = rsIcube.Fields(targetFieldName).Type
            newValue = clsCom.GetDefaultValue(fieldType, oldValue)
            
            ' 値が変更される場合だけ更新
            If Nz(newValue, "") <> Nz(oldValue, "") Then
                rsIcube.Edit
                rsIcube.Fields(targetFieldName).Value = newValue
                rsIcube.Update
            End If
            
            rsIcube.MoveNext
        Loop
        
        rsIcube.Close
        Set rsIcube = Nothing
        
        rsTarget.MoveNext
    Loop

CleanUp:
    If Not rsTarget Is Nothing Then
        rsTarget.Close
        Set rsTarget = Nothing
    End If
    
    If Not rsIcube Is Nothing Then
        rsIcube.Close
        Set rsIcube = Nothing
    End If

    Set db = Nothing
    
    MsgBox "エラー値置換え処理が完了しました。", vbInformation
End Sub
