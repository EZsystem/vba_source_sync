Attribute VB_Name = "mod_icubeVal7"
'============================================
' プロシージャ名: Update_追加工事名称_Cle
' Module: acc_mod_追加工事名称編集
' 概要: Icube_テーブルの「追加工事名称」から「追加工事名称_cle」に整形転写する
'       マスタ定義に基づく前処理後に、空白削除・括弧削除・全角化・㈱→(株)置換を行う
'============================================
Public Sub Update_追加工事名称_Cle()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim originalStr As String
    Dim baseStr As String
    Dim cleanedStr As String

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT [No], [発注者コード], [追加工事名称], [追加工事名称_cle] FROM Icube_", dbOpenDynaset)

    Do Until rs.EOF
        originalStr = Nz(rs!追加工事名称, "")
        
        ' --- 0. マスタベースで前処理した名称を取得 ---
        baseStr = GetCleanedName_FromMaster(Nz(rs!発注者コード, ""), originalStr)

        ' --- 1. ブランク削除（全角・半角スペース、タブ） ---
        cleanedStr = replace(baseStr, " ", "")
        cleanedStr = replace(cleanedStr, "　", "")
        cleanedStr = replace(cleanedStr, vbTab, "")

        ' --- 2. 【】内の文字ごと削除 ---
        Do While InStr(cleanedStr, "【") > 0 And InStr(cleanedStr, "】") > InStr(cleanedStr, "【")
            cleanedStr = Left(cleanedStr, InStr(cleanedStr, "【") - 1) & Mid(cleanedStr, InStr(cleanedStr, "】") + 1)
        Loop

        ' --- 3. 半角→全角（英数字・記号・カナなど） ---
        cleanedStr = StrConv(cleanedStr, vbWide)

        ' --- 4. ㈱ → (株) に置換（会社表記の簡略） ---
        cleanedStr = replace(cleanedStr, "㈱", "(株)")

        ' --- 5. 結果を転写 ---
        rs.Edit
        rs!追加工事名称_cle = cleanedStr
        rs.Update

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    'MsgBox "追加工事名称_cle を更新しましたニャー！", vbInformation

End Sub

