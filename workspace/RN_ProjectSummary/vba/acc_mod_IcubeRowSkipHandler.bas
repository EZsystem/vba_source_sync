Attribute VB_Name = "mod_IcubeRowSkipHandler"
'-------------------------------------
' Module: mod_IcubeRowSkipHandler
' 説明　：tbl_Temp_Icube_Import のレコード転写スキップ処理
' 作成日：2025/05/08
' 更新日：-
'-------------------------------------
Option Compare Database
Option Explicit

'=================================================
' サブルーチン名 : ApplyIcubeRowSkip
' 説明   : マスタ(tbl_xl_IcubeRowSkip)の設定に従い、
'         tbl_Temp_Icube_Import から該当レコードを削除する
' 引数   : なし
'=================================================
Public Sub ApplyIcubeRowSkip()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim skipField As String
    Dim skipVal As Variant
    Dim sqlDel As String

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_xl_IcubeRowSkip", dbOpenSnapshot)

    Do Until rs.EOF
        skipField = rs!対象フィールド名
        skipVal = rs!削除対象値

        ' DELETE 文を組み立て
        sqlDel = "DELETE FROM [tbl_Temp_Icube_Import] " & _
                 "WHERE [" & skipField & "] = " & IIf( _
                     IsNumeric(skipVal), _
                     skipVal, _
                     "'" & Replace(CStr(skipVal), "'", "''") & "'" _
                 )

        db.Execute sqlDel, dbFailOnError
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


