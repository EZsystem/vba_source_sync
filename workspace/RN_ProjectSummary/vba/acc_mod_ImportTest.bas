Attribute VB_Name = "acc_mod_ImportTest"
'Attribute VB_Name = "acc_mod_ImportTest"
Option Compare Database
Option Explicit

'============================================
' プロシージャ名 : Execute_All_Genka_Process
' 概要          : インポートからデータ加工(Test2, Test3)までの一連の処理を
'               順次実行し、エラーが発生した場合はそこで安全に停止する。
'============================================
Public Sub Execute_All_Genka_Process()
    Debug.Print "全体の処理を開始します..."
    
    If Not Import_Genka_Raw_Test() Then
        MsgBox "インポート処理でエラーが発生したため、後続のデータ加工を中止します。", vbCritical
        Exit Sub
    End If
    
    If Not Run_Genka_Normalization_Logic() Then
        MsgBox "正規化処理でエラーが発生したため、残りの処理を中止します。", vbCritical
        Exit Sub
    End If
    
    If Not Extract_Test3_F2_Values() Then
        MsgBox "抽出処理でエラーが発生しました。", vbCritical
        Exit Sub
    End If
    
    ' 追加: テストテーブルから本番テーブルへの転写・属性補正処理をシームレスに呼び出し
    Debug.Print "=== 本番テーブルへのデータ転写を開始します ==="
    Call acc_mod_Genka_Main.Import_GenkaData_ToMain
    
    ' ※Import_GenkaData_ToMain内部で完了通知が実行されるため、ここでのポップアップは不要
End Sub


'============================================
' プロシージャ名 : Import_Genka_Raw_Test
' 概要          : A7からA列の最終データ行までを自動取得して取り込む
'============================================
Public Function Import_Genka_Raw_Test() As Boolean
    Import_Genka_Raw_Test = False
    Dim db As DAO.Database
    Dim xlPath As String
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim lastRow As Long
    Dim importRange As String
    
    On Error GoTo Err_Handler
    
    Set db = CurrentDb
    xlPath = "D:\My_code\11_workspaces\RN_kanri_system\genka_system\原価システムimport.xlsm"
    
    ' 1. Excelを裏で開いてA列の最終行を取得する(Late Binding)
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(xlPath, ReadOnly:=True)
    Set xlSheet = xlBook.Sheets("原価S直データ")
    
    ' 行末から上へ探索して最終データ行を取得 (xlUp = -4162)
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row
    
    ' 取得したらExcelプロセスを終了
    xlBook.Close False
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    ' データが7行目以降に無い場合は終了
    If lastRow < 7 Then
        Debug.Print "A7行以降にデータが存在しません。"
        Exit Function
    End If
    
    ' 2. テストテーブルを空にする
    db.Execute "DELETE FROM [at_Test_Raw_Import]", dbFailOnError
    
    ' 3. 動的に範囲指定文字列を作成してインポート
    importRange = "原価S直データ!A7:AT" & lastRow
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, _
        "at_Test_Raw_Import", xlPath, False, importRange
        
    Debug.Print "--- インポート完了 ---"
    Debug.Print "テーブル [at_Test_Raw_Import] に " & (lastRow - 6) & " 行取り込みました。"
    Import_Genka_Raw_Test = True

Exit_Sub:
    ' 強制終了等の場合に残ったプロセスを解放
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Debug.Print "エラー (" & Err.Number & "): " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "インポートエラー"
    Import_Genka_Raw_Test = False
    Resume Exit_Sub
End Function


'============================================
' プロシージャ名 : Run_Genka_Normalization_Logic
' 概要          : 入力テーブル(at_Test_Raw_Import)のF1フィールドの値を元に、
'               1, 2, 3の行を1グループとしてデータを再構成し、
'               出力テーブル(at_Test2_Raw_Import)に転写する。
'============================================
Public Function Run_Genka_Normalization_Logic() As Boolean
    Run_Genka_Normalization_Logic = False
    Dim db As DAO.Database
    Dim rsIn As DAO.Recordset
    Dim rsOut As DAO.Recordset
    Dim rec As Object ' Late binding for Scripting.Dictionary
    Dim cleaner As New acc_clsDataCleaner_Test
    Dim i As Integer
    Dim f1_val As String
    Dim f3_val As String
    Dim spacePos As Integer

    On Error GoTo Err_Handler

    Set db = CurrentDb

    ' 出力先テーブルをクリア
    db.Execute "DELETE FROM [at_Test2_Raw_Import]", dbFailOnError

    Set rsIn = db.OpenRecordset("at_Test_Raw_Import", dbOpenSnapshot)
    Set rsOut = db.OpenRecordset("at_Test2_Raw_Import", dbOpenDynaset)

    Debug.Print "=== F1の値に基づいた正規化処理を開始します ==="
    Do Until rsIn.EOF
        f1_val = cleaner.CleanText(rsIn.Fields("F1").Value)

        If f1_val = "1" Then
            ' --- ステップ1: ヘッダー行(F1="1") を発見 ---
            Set rec = CreateObject("Scripting.Dictionary")
            Debug.Print "グループ化処理開始 (入力元 " & rsIn.AbsolutePosition + 1 & "行目)"

            ' F1="1"の行から全フィールドの初期値をコピー
            For i = 1 To 46
                rec("F" & i) = rsIn.Fields("F" & i).Value
            Next i

            ' F3を分割して新しいF1とF3の値を設定
            f3_val = cleaner.CleanText(rsIn.Fields("F3").Value)
            spacePos = InStr(f3_val, " ")

            If spacePos > 0 Then
                rec("F1") = Trim(Left(f3_val, spacePos - 1))
                rec("F3") = Trim(Mid(f3_val, spacePos + 1))
            Else
                rec("F1") = f3_val ' スペースがない場合はF3全体をF1へ
                rec("F3") = ""
            End If

            rsIn.MoveNext

            ' --- ステップ2: F1=2, F1=3 の行をチェック ---
            If Not rsIn.EOF Then
                f1_val = cleaner.CleanText(rsIn.Fields("F1").Value)
                If f1_val = "2" Then
                    Debug.Print "  - F1=2 のデータで上書き"
                    For i = 1 To 46
                        If i <> 1 And i <> 3 Then
                           If Not IsNull(rsIn.Fields("F" & i).Value) And Trim(rsIn.Fields("F" & i).Value & "") <> "" Then
                                rec("F" & i) = rsIn.Fields("F" & i).Value
                           End If
                        End If
                    Next i
                    rsIn.MoveNext

                    If Not rsIn.EOF Then
                         f1_val = cleaner.CleanText(rsIn.Fields("F1").Value)
                         If f1_val = "3" Then
                             Debug.Print "  - F1=3 のデータで上書き"
                             For i = 1 To 46
                                 If i <> 1 And i <> 3 Then
                                    If Not IsNull(rsIn.Fields("F" & i).Value) And Trim(rsIn.Fields("F" & i).Value & "") <> "" Then
                                        rec("F" & i) = rsIn.Fields("F" & i).Value
                                    End If
                                 End If
                             Next i
                             rsIn.MoveNext
                         End If
                    End If
                End If
            End If
            ' --- ステップ3: 構築したレコードを書き込み ---
            rsOut.AddNew
            For i = 1 To 46
                rsOut.Fields("F" & i) = rec("F" & i)
            Next i
            rsOut.Update
            Debug.Print "  - グループ化処理完了、at_Test2_Raw_Importにレコードを書き込みました。"

            Set rec = Nothing
        Else
            rsIn.MoveNext
        End If
    Loop
    Debug.Print "=== F1基準の正規化処理が正常に完了しました ==="
    Run_Genka_Normalization_Logic = True

Exit_Sub:
    If Not rsIn Is Nothing Then rsIn.Close
    If Not rsOut Is Nothing Then rsOut.Close
    Set rsIn = Nothing
    Set rsOut = Nothing
    Set db = Nothing
    Set rec = Nothing
    Exit Function

Err_Handler:
    Debug.Print "エラー (" & Err.Number & "): " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "正規化処理エラー"
    Run_Genka_Normalization_Logic = False
    Resume Exit_Sub
End Function


'============================================
' プロシージャ名 : Extract_Test3_F2_Values
' 概要          : 入力テーブル(at_Test_Raw_Import)を走査し、
'               F2フィールドに値がある行を抽出して、
'               出力テーブル(at_Test3_Raw_Import)に転写する。
'============================================
Public Function Extract_Test3_F2_Values() As Boolean
    Extract_Test3_F2_Values = False
    Dim db As DAO.Database
    Dim rsIn As DAO.Recordset
    Dim rsOut3 As DAO.Recordset
    Dim i As Integer
    Dim f2_val As String

    On Error GoTo Err_Handler

    Set db = CurrentDb

    ' 出力先テーブルをクリア
    db.Execute "DELETE FROM [at_Test3_Raw_Import]", dbFailOnError

    Set rsIn = db.OpenRecordset("at_Test_Raw_Import", dbOpenSnapshot)
    Set rsOut3 = db.OpenRecordset("at_Test3_Raw_Import", dbOpenDynaset)

    Debug.Print "=== F2値に基づくテスト抽出処理を開始します ==="
    Do Until rsIn.EOF
        ' F2の値をチェック：空文字でなく、かつ「管理番号」でもない場合
        f2_val = Trim(rsIn.Fields("F2").Value & "")
        If f2_val <> "" And f2_val <> "管理番号" Then
            rsOut3.AddNew
            For i = 1 To 46
                rsOut3.Fields("F" & i) = rsIn.Fields("F" & i).Value
            Next i
            rsOut3.Update
        End If
        rsIn.MoveNext
    Loop
    Debug.Print "=== F2値に基づく抽出処理が完了しました ==="
    Extract_Test3_F2_Values = True

Exit_Sub:
    If Not rsIn Is Nothing Then rsIn.Close
    If Not rsOut3 Is Nothing Then rsOut3.Close
    Set rsIn = Nothing
    Set rsOut3 = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Debug.Print "エラー (" & Err.Number & "): " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "F2抽出処理エラー"
    Extract_Test3_F2_Values = False
    Resume Exit_Sub
End Function




