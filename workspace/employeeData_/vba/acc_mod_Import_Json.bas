Attribute VB_Name = "acc_mod_Import_Json"
Public Function ImportJsonProcess()
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim rawText As String
    Dim db As DAO.Database
    Dim rows() As String
    Dim i As Long
    Dim recordCount As Long
    
    filePath = "D:\My_code\11_workspaces\RN_kanri_system\kenmu_system\data_to_access.json"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' --- 経過表示1：ファイル確認 ---
    If Not fso.FileExists(filePath) Then
        MsgBox "JSONファイルが見つかりません: " & filePath
        Exit Function
    End If

    ' 1. JSON読み込み
    Set ts = fso.OpenTextFile(filePath, 1, False)
    rawText = ts.ReadAll
    ts.Close

    ' --- 経過表示2：読み込んだ文字数を確認 ---
    If MsgBox("JSONを読み込みました。文字数: " & Len(rawText) & " 文字。中身を表示しますか？", vbYesNo) = vbYes Then
        MsgBox Left(rawText, 500) ' 最初の500文字を表示
    End If

    ' 2. テーブルクリア
    Set db = CurrentDb
    db.Execute "DELETE * FROM at_testTemp;"
    Debug.Print "テーブルをクリアしました。"

    ' 3. 解析とインポート
    rows = Split(rawText, "}")
    recordCount = 0
    
    ' ※エラーを隠さないために On Error Resume Next は使いません
    For i = LBound(rows) To UBound(rows) - 1
        Dim val_no As String
        val_no = GetJsonKeyValue(rows(i), "社員番号")
        
        If val_no <> "" Then
            Dim sql As String
            sql = "INSERT INTO at_testTemp (社員番号, 氏名_戸籍上, 氏名カナ, 氏名_ﾒｰﾙ表示用, 資格, 所属, 役職, 対外呼称) " & _
                  "VALUES ('" & val_no & "', '" & _
                  GetJsonKeyValue(rows(i), "氏名_戸籍上") & "', '" & _
                  GetJsonKeyValue(rows(i), "氏名カナ") & "', '" & _
                  GetJsonKeyValue(rows(i), "氏名_ﾒｰﾙ表示用") & "', '" & _
                  GetJsonKeyValue(rows(i), "資格") & "', '" & _
                  GetJsonKeyValue(rows(i), "所属") & "', '" & _
                  GetJsonKeyValue(rows(i), "役職") & "', '" & _
                  GetJsonKeyValue(rows(i), "対外呼称") & "');"
            
            db.Execute sql
            recordCount = recordCount + 1
        End If
    Next i

    ' --- 経過表示3：最終結果 ---
    MsgBox "処理終了。インポート件数: " & recordCount & " 件"
    
    ' テストのため DoCmd.Quit はコメントアウト（手動で閉じてください）
    ' DoCmd.Quit
End Function

Private Function GetJsonKeyValue(ByVal txt As String, ByVal key As String) As String
    Dim s As String
    On Error GoTo ErrHand
    s = Split(txt, """" & key & """: """)(1)
    GetJsonKeyValue = Split(s, """")(0)
    Exit Function
ErrHand:
    GetJsonKeyValue = ""
End Function

