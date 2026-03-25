Attribute VB_Name = "mod_icubeVal22"
'-------------------------------------
' Module: acc_mod_IcubeOrderMapping
' 説明  : Icube_テーブルの受注用フィールドに仮基本工事情報を転写する処理群
' 作成日: 2025/07/03
' 更新日: -
'-------------------------------------
Option Compare Database
Option Explicit

'============================================================
' プロシージャ名: MapOrderFieldsToIcube
' 概要         : Icube_テーブルで一件工事判定が「小口工事」の行を対象に、
'                tb_仮基本工事の情報を「作業所」「Q」「官民」キーで一致した場合、
'                仮基本工事コード・略名をIcube_に転写する
'============================================================
Public Sub MapOrderFieldsToIcube()
    On Error GoTo Err_Handler

    ' --- 1. 初期化 ---
    Dim db As DAO.Database
    Dim rsIcube As DAO.Recordset
    Dim rsRef As DAO.Recordset
    Dim sqlIcube As String, sqlRef As String
    Dim workSite As String, qStr As String, publicPrivate As String
    Dim matchCount As Long
    Dim cleaner As acc_clsDataCleaner
    Set cleaner = New acc_clsDataCleaner

    Set db = CurrentDb

    sqlRef = "SELECT * FROM tb_仮基本工事"
    Set rsRef = db.OpenRecordset(sqlRef, dbOpenSnapshot)

    sqlIcube = "SELECT * FROM Icube_ WHERE 一件工事判定 = '小口工事'"
    Set rsIcube = db.OpenRecordset(sqlIcube, dbOpenDynaset)

    matchCount = 0

    ' --- 2. 主処理 ---
    Do While Not rsIcube.EOF
        Dim isHit As Boolean
        isHit = False

        ' フィールドの値をクレンジングして取得
        workSite = cleaner.CleanText(rsIcube!基本工事名_作業所)
        publicPrivate = cleaner.CleanText(rsIcube!基本工事名_官民)
        qStr = cleaner.CleanText(rsIcube!受注Q)
        qStr = ConvertToZenkakuNumber(qStr) & "Q"

        rsRef.MoveFirst
        Do While Not rsRef.EOF
            If cleaner.CleanText(rsRef!基本工事名_作業所) = workSite _
               And cleaner.CleanText(rsRef!基本工事名_Q) = qStr _
               And cleaner.CleanText(rsRef!基本工事名_官民) = publicPrivate Then

                rsIcube.Edit
                rsIcube!仮基本工事コード_受注 = rsRef!仮基本工事コード
                rsIcube!仮基本工事略名_受注 = rsRef!仮基本工事略名
                rsIcube.Update

                matchCount = matchCount + 1
                isHit = True
                Exit Do
            End If
            rsRef.MoveNext
        Loop

        rsIcube.MoveNext
    Loop

    ' --- 3. 結果表示（必要に応じてコメントアウト解除） ---
    'MsgBox "受注用の更新完了" & vbCrLf & "一致して転写された件数：" & matchCount, vbInformation

Exit_Handler:
    On Error Resume Next
    rsIcube.Close
    rsRef.Close
    Set rsIcube = Nothing
    Set rsRef = Nothing
    Set db = Nothing
    Set cleaner = Nothing
    Exit Sub

Err_Handler:
    MsgBox "エラーが発生しました: " & Err.description, vbExclamation
    Resume Exit_Handler
End Sub

'============================================================
' 関数名: ConvertToZenkakuNumber
' 概要  : 半角数字文字列を全角数字へ変換して返す
' 引数  : halfNumStr - 半角数字文字列
' 戻り値: 全角数字文字列
'============================================================
Private Function ConvertToZenkakuNumber(ByVal halfNumStr As String) As String
    Dim i As Integer
    Dim result As String
    result = ""

    For i = 1 To Len(halfNumStr)
        Dim ch As String
        ch = Mid(halfNumStr, i, 1)
        Select Case ch
            Case "0": result = result & "０"
            Case "1": result = result & "１"
            Case "2": result = result & "２"
            Case "3": result = result & "３"
            Case "4": result = result & "４"
            Case "5": result = result & "５"
            Case "6": result = result & "６"
            Case "7": result = result & "７"
            Case "8": result = result & "８"
            Case "9": result = result & "９"
            Case Else: result = result & ch
        End Select
    Next i

    ConvertToZenkakuNumber = result
End Function

'-------------------------------------
'（モジュール終わり）
'-------------------------------------


