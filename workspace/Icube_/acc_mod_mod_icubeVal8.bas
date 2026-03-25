Attribute VB_Name = "mod_icubeVal8"
'-------------------------------------
' Module: mod_icubeVal8
' 説明　：tbl_顧客データ → Icube_ の条件付き転写処理
'         顧客コード＝発注者コードが一致かつ未転写の場合のみ処理する
' 作成日：2025/07/30
' 更新日：-
'-------------------------------------
Option Compare Database
Option Explicit

'============================================
' プロシージャ名: Transfer_顧客名_IfNotExists
' Module       : mod_icubeVal8
' 概要         : 仮テーブルの会社名を本テーブルの発注者名_tblへ転写
'                一致する発注者コードが存在かつ、未転写時のみ処理する
'============================================
Public Sub Transfer_顧客名_IfNotExists()

    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim strSQL As String
    Dim customerCode As String
    Dim customerName As String

    Set db = CurrentDb
    Set rsSource = db.OpenRecordset("tbl_顧客データ", dbOpenSnapshot)

    Do Until rsSource.EOF
        customerCode = Nz(rsSource!顧客コード, "")
        customerName = Nz(rsSource!会社名, "")

        If customerCode <> "" And customerName <> "" Then
            strSQL = "SELECT 発注者名_tbl FROM Icube_ " & _
                     "WHERE 発注者コード = '" & customerCode & "' " & _
                     "AND (発注者名_tbl IS NULL OR 発注者名_tbl = '')"

            Set rsTarget = db.OpenRecordset(strSQL, dbOpenDynaset)

            Do Until rsTarget.EOF
                rsTarget.Edit
                rsTarget!発注者名_tbl = customerName
                rsTarget.Update
                rsTarget.MoveNext
            Loop

            rsTarget.Close
            Set rsTarget = Nothing
        End If

        rsSource.MoveNext
    Loop

    rsSource.Close
    Set rsSource = Nothing
    Set db = Nothing

    'MsgBox "顧客データの転写が完了しましたニャー！", vbInformation

End Sub


