Attribute VB_Name = "mod_icubeVal2"


Option Compare Database
Option Explicit

'テーブルIcubeへ基本工事名称を分割して記入
Public Sub mod_icube_Val2ALL()

'テーブルIcubeへ基本工事名称の分割記入
    Call mod_icube_Val2copy1
'テーブルIcubeへ基本工事名称からの期記入
    Call mod_icube_Val2copy2
'テーブルIcubeへ仮基本工事コード転写
    Call mod_icube_Val2copy3
'テーブルIcubeへ仮基本工事コード受注　転写
'別モジュール：mod_icubeVal22
    Call MapOrderFieldsToIcube

End Sub



'============================================
' プロシージャ名 : mod_icube_Val2copy3_Updated
' モジュール名   : xl_mod_CostMng1（など任意）
' 概要 : Icube_ に仮基本工事コードと略名を転写（小口工事対象）
'============================================
Public Sub mod_icube_Val2copy3()
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rsIcube As DAO.Recordset
    Dim rsRef As DAO.Recordset
    Dim sqlIcube As String, sqlRef As String
    Dim 作業所 As String, Q As String, 官民 As String, 繰越 As String
    Dim matchCount As Long

    Set db = CurrentDb

    ' 参照テーブル（仮基本工事マスター）
    sqlRef = "SELECT * FROM tb_仮基本工事"
    Set rsRef = db.OpenRecordset(sqlRef, dbOpenSnapshot)

    ' Icube_（更新対象）で「一件工事判定 = '小口工事'」のみ
    sqlIcube = "SELECT * FROM Icube_ WHERE 一件工事判定 = '小口工事'"
    Set rsIcube = db.OpenRecordset(sqlIcube, dbOpenDynaset)

    matchCount = 0

    Do While Not rsIcube.EOF
        Dim hitFound As Boolean
        hitFound = False

        作業所 = Trim(Nz(rsIcube!基本工事名_作業所, ""))
        Q = Trim(Nz(rsIcube!基本工事名_Q, ""))
        官民 = Trim(Nz(rsIcube!基本工事名_官民, ""))
        繰越 = Trim(Nz(rsIcube!基本工事名_繰越, ""))

        rsRef.MoveFirst
        Do While Not rsRef.EOF
            If Trim(Nz(rsRef!基本工事名_作業所, "")) = 作業所 And _
               Trim(Nz(rsRef!基本工事名_Q, "")) = Q And _
               Trim(Nz(rsRef!基本工事名_官民, "")) = 官民 And _
               Trim(Nz(rsRef!基本工事名_繰越, "")) = 繰越 Then

                rsIcube.Edit
                rsIcube!仮基本工事コード = rsRef!仮基本工事コード
                rsIcube!仮基本工事略名 = rsRef!仮基本工事略名  ' ← 追加転写
                rsIcube.Update

                matchCount = matchCount + 1
                hitFound = True
                Exit Do
            End If
            rsRef.MoveNext
        Loop

        rsIcube.MoveNext
    Loop

    'MsgBox "更新完了ニャ！" & vbCrLf & "一致して転写された件数：" & matchCount, vbInformation

Exit_Handler:
    On Error Resume Next
    rsIcube.Close
    rsRef.Close
    Set rsIcube = Nothing
    Set rsRef = Nothing
    Set db = Nothing
    Exit Sub

Err_Handler:
    MsgBox "エラーが発生したにゃ：" & Err.description, vbExclamation
    Resume Exit_Handler
End Sub



'テーブルIcubeへ基本工事名称からの期記入
Sub mod_icube_Val2copy2()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim n年度 As Variant
    Dim 元文字列 As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Icube_", dbOpenDynaset)
    
    Do While Not rs.EOF
        元文字列 = Nz(rs!基本工事名_年度, "")
        
        If 元文字列 <> "" And 元文字列 <> "N/A" Then
            ' 全角を半角に変換してから数値化
            n年度 = val(StrConv(元文字列, vbNarrow)) - 12
        Else
            n年度 = Null
        End If
        
        rs.Edit
        rs!基本工事名_期 = n年度
        rs.Update
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    'MsgBox "値が全部 ?12 だった問題、解消したはずニャ?！", vbInformation
End Sub



'テーブルIcubeへ基本工事名称の分割記入
Sub mod_icube_Val2copy1()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim 元文字列 As String
    Dim pos年度 As Long, posRN As Long, 作業所 As String
    Dim 年度 As String, Q As String, 官民 As String, 繰越 As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Icube_", dbOpenDynaset)
    
    Do While Not rs.EOF
        If rs!一件工事判定 = "小口工事" Then
            元文字列 = Nz(rs!基本工事名称, "")
            
            '◆作業所
            If Left(元文字列, 3) = "建築部" Then
                作業所 = "建築部"
            Else
                posRN = InStr(元文字列, "ＲＮ")
                If posRN >= 3 Then
                    作業所 = Mid(元文字列, posRN - 2, 2)
                Else
                    作業所 = ""
                End If
            End If
            
            '◆年度
            pos年度 = InStr(元文字列, "年度")
            If pos年度 >= 3 Then
                年度 = Mid(元文字列, pos年度 - 2, 2)
            Else
                年度 = ""
            End If
            
            '◆Q
            If pos年度 > 0 And Len(元文字列) >= pos年度 + 2 Then
                Q = Mid(元文字列, pos年度 + 2, 2)
            Else
                Q = ""
            End If
            
            '◆官民
            If InStr(元文字列, "民間") > 0 Then
                官民 = "民間"
            ElseIf InStr(元文字列, "官庁") > 0 Then
                官民 = "官庁"
            Else
                官民 = ""
            End If
            
            '◆繰越
            If InStr(元文字列, "（繰越）") > 0 Then
                繰越 = "（繰越）"
            Else
                繰越 = ""
            End If
            
        Else
            ' 一件工事判定が小口工事以外
            作業所 = "N/A"
            年度 = "N/A"
            Q = "N/A"
            官民 = "N/A"
            繰越 = "N/A"
        End If
        
        ' 転写
        rs.Edit
        rs!基本工事名_作業所 = 作業所
        rs!基本工事名_年度 = 年度
        rs!基本工事名_Q = Q
        rs!基本工事名_官民 = 官民
        rs!基本工事名_繰越 = 繰越
        rs.Update
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    'MsgBox "転写処理が完了しましたニャ?！", vbInformation
End Sub


