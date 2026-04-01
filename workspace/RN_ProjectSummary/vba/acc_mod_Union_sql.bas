Attribute VB_Name = "acc_mod_Union_sql"
' Module: acc_mod_Union_sql
'============================================================
' サブルーチン名 : Transfer_Union_NoDuplicate
' 概要           : クエリから s基本工事コード ごとに重複を排除して転写
' 依存関係       : acc_mod_MappingTemplate (テーブル名・クエリ名定数)
'============================================================
Public Sub Transfer_Union_NoDuplicate()
    On Error GoTo ErrHandler

    Dim db As DAO.Database         ' データベースオブジェクト
    Dim rs As DAO.Recordset        ' クエリのレコードセット
    Dim dict As Object             ' s基本工事コード重複チェック用Dictionary
    Dim sql As String              ' SQL文格納用文字列

    Set db = CurrentDb
    Set dict = CreateObject("Scripting.Dictionary")

    ' 1. 転写先テーブルを事前クリア (定数 AT_LINK_KIHON_NAME を使用)
    db.Execute "DELETE FROM [" & AT_LINK_KIHON_NAME & "]", dbFailOnError

    ' 2. クエリのSQLを実行 (定数 AQ_SEL_KIHON_NAME を使用)
    ' クエリ名は名称変更対象外ですが、定数管理に合わせることで保守性を高めています
    sql = "SELECT s基本工事コード, s基本工事名称, 完工期, 完工Q, 施工管轄組織名, 一件工事判定 FROM [" & AQ_SEL_KIHON_NAME & "]"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    ' 3. クエリ結果を1行ずつ処理
    Do While Not rs.EOF
        Dim empCode As String
        empCode = rs!s基本工事コード

        ' Dictionaryで未登録のコードのみ転写実行
        If Not dict.Exists(empCode) Then
            ' INSERT先のテーブル名を定数に変更
            db.Execute "INSERT INTO [" & AT_LINK_KIHON_NAME & "] " & _
                       "(s基本工事コード, s基本工事名称, 完工期, 完工Q, 施工管轄組織名, 一件工事判定) " & _
                       "VALUES (" & _
                       "'" & Replace(empCode, "'", "''") & "', " & _
                       "'" & Replace(Nz(rs!s基本工事名称, ""), "'", "''") & "', " & _
                       "'" & Replace(Nz(rs!完工期, ""), "'", "''") & "', " & _
                       "'" & Replace(Nz(rs!完工Q, ""), "'", "''") & "', " & _
                       "'" & Replace(Nz(rs!施工管轄組織名, ""), "'", "''") & "', " & _
                       "'" & Replace(Nz(rs!一件工事判定, ""), "'", "''") & "')", dbFailOnError

            dict.Add empCode, True   ' 重複登録防止用に登録
        End If
        rs.MoveNext
    Loop

    ' 4. 解放処理
    rs.Close: Set rs = Nothing
    Set dict = Nothing: Set db = Nothing

    MsgBox "重複排除して転写完了したニャ", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生したにゃ：" & Err.Description, vbCritical
    If Not rs Is Nothing Then rs.Close
End Sub

