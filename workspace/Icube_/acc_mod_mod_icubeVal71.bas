Attribute VB_Name = "mod_icubeVal71"
'-------------------------------------
' Module: mod_icubeNameCleaner
' 説明　：tbl_工事名cle を元に、Icube_累計 の追加工事名称_cle を生成
'         条件に応じて先頭ワード削除・トリガーマッチング処理を実施
' 作成日：2025/07/30
' 更新日：-
'-------------------------------------
Option Compare Database
Option Explicit


'============================================
' 関数名 : GetCleanedName_FromMaster
' 概要   : 発注者コードと元名称に基づき、マスタ条件で削除語等を適用した初期値を返す
' 引数   : 発注者コード（String）, 元の追加工事名称（String）
' 戻り値 : 加工済みの中間名称（String）
'============================================
Public Function GetCleanedName_FromMaster(発注者コード As String, 元名称 As String) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim トリガーワード As String
    Dim delWord As String
    Dim result As String
    Dim posDel As Long

    result = 元名称
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_工事名cle", dbOpenSnapshot)

    Do Until rs.EOF
        If Nz(rs!発注者コード, "") = 発注者コード Then
            トリガーワード = Nz(rs!トリガーワード, "")
            delWord = Nz(rs!del区分ワード, "")

            ' --- トリガーワードが空、または冒頭一致のときのみ処理 ---
            If トリガーワード = "" Or Left(result, Len(トリガーワード)) = トリガーワード Then

                If delWord = "ブランク" Then
                    ' --- 最初のスペースを基準に右側を残す ---
                    posDel = InStr(result, " ")
                    If posDel = 0 Then posDel = InStr(result, "　") ' 全角スペース
                    If posDel > 0 Then
                        result = Mid(result, posDel + 1)
                        result = Trim(result)
                    End If
                Else
                    ' --- delWord が文字列で見つかった場合、その右側を残す ---
                    posDel = InStr(result, delWord)
                    If posDel > 0 Then
                        result = Mid(result, posDel + Len(delWord))
                        result = Trim(result)
                    End If
                End If

                Exit Do ' 最初に一致した条件のみ適用
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetCleanedName_FromMaster = result
End Function



'============================================
' プロシージャ名: Generate_追加工事名称_cle_FromMaster
' Module       : mod_icubeVal71
' 概要         : 仮テーブルと本テーブルを突合し、
'                トリガーワードと削除条件に基づいて追加工事名称_cleを整形転写する
'============================================
Public Sub Generate_追加工事名称_cle_FromMaster()
    
    Dim db As DAO.Database
    Dim rsMaster As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim sql As String
    Dim 発注者コード As String
    Dim トリガーワード As String
    Dim delWord As String
    Dim 追加工事名称 As String
    Dim cleanedName As String

    Set db = CurrentDb
    Set rsMaster = db.OpenRecordset("tbl_工事名cle", dbOpenSnapshot)

    Do Until rsMaster.EOF
        発注者コード = Nz(rsMaster!発注者コード, "")
        トリガーワード = Nz(rsMaster!トリガーワード, "")
        delWord = Nz(rsMaster!del区分ワード, "")
        
        ' ブランク指定時の処理
        If delWord = "ブランク" Then
            delWord = "　" ' 全角スペース
        End If

        If 発注者コード <> "" And トリガーワード <> "" Then
            sql = "SELECT No, 追加工事名称 FROM Icube_累計 " & _
                  "WHERE 発注者コード = '" & 発注者コード & "' " & _
                  "AND 追加工事名称 LIKE '" & トリガーワード & "%'"

            Set rsTarget = db.OpenRecordset(sql, dbOpenDynaset)

            Do Until rsTarget.EOF
                追加工事名称 = Nz(rsTarget!追加工事名称, "")

                ' 初期値：元のまま
                cleanedName = 追加工事名称

                ' 削除ワードが先頭一致していれば削除
                If Left(追加工事名称, Len(delWord)) = delWord Then
                    cleanedName = Mid(追加工事名称, Len(delWord) + 1)
                    cleanedName = Trim(cleanedName)
                End If

                ' 転写処理
                rsTarget.Edit
                rsTarget!追加工事名称_cle = cleanedName
                rsTarget.Update

                rsTarget.MoveNext
            Loop

            rsTarget.Close
            Set rsTarget = Nothing
        End If

        rsMaster.MoveNext
    Loop

    rsMaster.Close
    Set rsMaster = Nothing
    Set db = Nothing

    MsgBox "追加工事名称_cle の転写が完了したニャー！", vbInformation

End Sub


