Attribute VB_Name = "Union_sql"
'============================================================
' サブルーチン名 : Transfer_Union_NoDuplicate
' 概要           : クエリ sel_基本工事名称 の結果から s基本工事コード ごとに
'                  重複を排除して tblRink_基本工事名称 に転写する処理にゃ
' 処理内容       : Dictionaryを使って s基本工事コード をキーに一意判定
'                  最初の1件だけINSERTする構成にゃ
' 引数           : なし
' 戻り値         : なし
'============================================================
Public Sub Transfer_Union_NoDuplicate()
    On Error GoTo ErrHandler

    Dim db As DAO.Database         ' データベースオブジェクト
    Dim rs As DAO.Recordset        ' クエリのレコードセット
    Dim dict As Object             ' s基本工事コード重複チェック用Dictionary
    Dim sql As String              ' SQL文格納用文字列

    Set db = CurrentDb
    Set dict = CreateObject("Scripting.Dictionary")

    ' 転写先テーブルを事前クリア
    db.Execute "DELETE FROM tblRink_基本工事名称", dbFailOnError

    ' クエリ sel_基本工事名称 のSQLを実行（保存済クエリを使う構成）
    sql = "SELECT s基本工事コード, s基本工事名称, 完工期, 完工Q, 施工管轄組織名, 一件工事判定 FROM sel_基本工事名称"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    ' クエリ結果を1行ずつ処理
    Do While Not rs.EOF
        Dim empCode As String
        empCode = rs!s基本工事コード

        ' Dictionaryで未登録のコードのみ転写実行
        If Not dict.Exists(empCode) Then
            db.Execute "INSERT INTO tblRink_基本工事名称 " & _
                       "(s基本工事コード, s基本工事名称, 完工期, 完工Q, 施工管轄組織名, 一件工事判定) " & _
                       "VALUES (" & _
                       "'" & replace(empCode, "'", "''") & "', " & _
                       "'" & replace(Nz(rs!s基本工事名称, ""), "'", "''") & "', " & _
                       "'" & replace(Nz(rs!完工期, ""), "'", "''") & "', " & _
                       "'" & replace(Nz(rs!完工Q, ""), "'", "''") & "', " & _
                       "'" & replace(Nz(rs!施工管轄組織名, ""), "'", "''") & "', " & _
                       "'" & replace(Nz(rs!一件工事判定, ""), "'", "''") & "')", dbFailOnError

            dict.Add empCode, True   ' 重複登録防止用に登録
        End If
        rs.MoveNext
    Loop

    ' 解放処理
    rs.Close: Set rs = Nothing
    Set dict = Nothing: Set db = Nothing

    MsgBox "重複排除して転写完了したニャ", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生したにゃ：" & Err.description, vbCritical
    If Not rs Is Nothing Then rs.Close
End Sub


