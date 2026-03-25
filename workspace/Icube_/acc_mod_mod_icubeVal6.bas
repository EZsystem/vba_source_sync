Attribute VB_Name = "mod_icubeVal6"
'-------------------------------------
' Module: acc_mod_MainUpdater
' 説明　：acc_clsFieldTranscriber クラスを使用して、条件付きフィールド転写を実行する
' 作成日：2025/05/12
'-------------------------------------
Option Compare Database
Option Explicit

'=================================================
' サブルーチン名 : Run_FieldTranscribe_WhenEmpty
' 説明   : 「s基本工事コード」が空 or "N/A" のレコードを対象に、
'          「工事コード」「工事帳票名」の値を
'          「s基本工事コード」「s基本工事名」へ転写するにゃ！
'=================================================
Public Sub Run_FieldTranscribe_WithSkipList()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cond As String

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM Icube_", dbOpenDynaset)

    Do While Not rs.EOF
        ' 条件フィールドの値をクリーンにして判定
        cond = Trim(UCase(Nz(rs!s基本工事コード, "")))

        If Not cond Like "KT*" Then
            rs.Edit
            rs!s基本工事コード = rs!工事コード
            rs!s基本工事名称 = rs!工事帳票名
            rs.Update
            'Debug.Print "転写:", rs!No, cond, "→", rs!工事コード
        End If

        rs.MoveNext
    Loop

    rs.Close
    'MsgBox "KT*を除いた転写が完了しましたにゃ！", vbInformation
End Sub

