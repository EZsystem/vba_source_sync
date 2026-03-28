Attribute VB_Name = "acc_mod_ImportTest"
Option Compare Database
Option Explicit

'============================================
' プロシージャ名 : Import_Genka_Raw_Test
' 概要          : A7:AT26 (20行分) を正確に取り込む
'============================================
Public Sub Import_Genka_Raw_Test()
    Dim db As DAO.Database
    Dim xlPath As String
    Set db = CurrentDb
    
    ' 1. テストテーブルを空にする
    db.Execute "DELETE FROM [at_Test_Raw_Import]", dbFailOnError
    
    xlPath = "D:\My_code\11_workspaces\RN_kanri_system\genka_system\原価システムimport.xlsm"
    
    ' 2. 範囲を A7:AT26 に限定してインポート
    ' (タイトル行なし False で取り込み)
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, _
        "at_Test_Raw_Import", xlPath, False, "原価S直データ!A7:AT26"
        
    Debug.Print "--- インポート完了 ---"
    Debug.Print "テーブル [at_Test_Raw_Import] に20行取り込みました。"
End Sub


'============================================
' プロシージャ名 : Debug_GenkaImport_Logic_v2
' 概要          : F1-F46の20行を1行ずつ詳細に実況解説する
'============================================
Public Sub Debug_GenkaImport_Logic_v2()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cleaner As New acc_clsDataCleaner_Test
    Dim currentKT As String
    Dim recordCount As Long
    
    Set db = CurrentDb
    ' インポートしたテーブルを確実に開く
    Set rs = db.OpenRecordset("at_Test_Raw_Import", dbOpenSnapshot)
    
    ' レコード数の確認
    If rs.EOF Then
        Debug.Print "【警告】テーブル [at_Test_Raw_Import] は空です。先に Import_Genka_Raw_Test を実行してください。"
        Exit Sub
    End If

    Debug.Print "--- 詳細検証(20行限定) 開始 ---"
    
    Do Until rs.EOF
        recordCount = recordCount + 1
        ' 1行ごとに生データを出力（F1=A列, F2=B列, F3=C列）
        Dim f1 As String: f1 = cleaner.CleanText(rs![f1])
        Dim f2 As String: f2 = cleaner.CleanText(rs![f2])
        Dim f3 As String: f3 = cleaner.CleanText(rs![f3])
        
        Debug.Print ">>> [Scan " & recordCount & "行目] RawData: F1=[" & f1 & "] F2=[" & f2 & "] F3=[" & f3 & "]"
        
        ' --- 判定ロジックの実況 ---
        
        ' 1. C列がKT番号か？ (親の判定)
        If f3 Like "KT*" Then
            currentKT = Left(f3, 9)
            Debug.Print "   => 【判定：親(基本)】管理番号 " & currentKT & " を保持しました。"
            
        ' 2. B列が数値か？ (子の判定)
        ElseIf IsNumeric(f2) And Len(f2) > 0 Then
            If currentKT = "" Then
                Debug.Print "   => 【判定：エラー】枝番 " & f2 & " が見つかりましたが、親(KT)がまだ出てきていません！"
            Else
                Debug.Print "   => 【判定：子(枝番)】枝番 " & f2 & " を「" & currentKT & "」に紐付け成功。"
                ' 金額列のテスト（K列=F11と想定）
                Debug.Print "      金額パーステスト(F11): " & cleaner.TextToCurrency(rs![F11])
            End If
            
        ' 3. 集計行などの無視判定
        ElseIf f2 Like "*合計*" Or f3 Like "*合計*" Or f2 Like "*経費*" Or f3 Like "*経費*" Then
            Debug.Print "   => 【判定：スキップ】集計用または不要な行として無視します。"
            
        ' 4. どれにも当てはまらない不明な行
        Else
            Debug.Print "   => 【判定：対象外】この行はシステム上不要な余白や不明なデータです。"
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Debug.Print "--- 検証終了：合計 " & recordCount & " 行をスキャンしました ---"
End Sub

'============================================
' プロシージャ名 : Run_Genka_Normalization_Logic
' 概要          : Raw(46列) のズレを修正し、Test2(46列) へ正規化して転写
'============================================
Public Sub Run_Genka_Normalization_Logic()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsIn As DAO.Recordset  ' 元ネタ：at_Test_Raw_Import
    Dim rsOut As DAO.Recordset ' 移動先：at_Test2_Raw_Import
    Dim cleaner As New acc_clsDataCleaner_Test
    
    Dim currentKT As String
    Dim parentName As String
    Dim i As Long, j As Integer

    ' 1. 移動先テーブルをリセット
    db.Execute "DELETE FROM [at_Test2_Raw_Import]", dbFailOnError
    
    Set rsIn = db.OpenRecordset("at_Test_Raw_Import", dbOpenSnapshot)
    Set rsOut = db.OpenRecordset("at_Test2_Raw_Import", dbOpenDynaset)
    
    Debug.Print "=== セル値の移動・正規化処理 開始 ==="

    Do Until rsIn.EOF
        i = i + 1
        Dim raw_F2 As String: raw_F2 = cleaner.CleanText(rsIn![f2])
        Dim raw_F3 As String: raw_F3 = cleaner.CleanText(rsIn![f3])
        
        ' --- A. 親（KT）行の情報を記憶 ---
        If raw_F3 Like "KT*" Then
            currentKT = Left(raw_F3, 9)
            parentName = Trim(Mid(raw_F3, 11))
            Debug.Print "Row " & i & ": 【親記憶】No:" & currentKT & " / Name:" & parentName
            
        ' --- B. 子（枝番）行の値を正しい列へ移動 ---
        ElseIf IsNumeric(raw_F2) And Len(raw_F2) > 0 Then
            rsOut.AddNew
            
            ' 【重要】ここが「セル値の移動」の定義です
            rsOut![F4] = parentName          ' 親の名前をセット
            rsOut![F5] = currentKT           ' 親の管理番号をセット
            rsOut![F6] = raw_F2              ' 枝番をセット
            rsOut![F7] = raw_F3              ' 子の工事名をセット
            rsOut![F8] = rsIn![F4]           ' 担当者名
            rsOut![F9] = rsIn![F5]           ' 状況
            rsOut![F10] = rsIn![F6]          ' 登録年月
            rsOut![F11] = rsIn![F7]          ' 工事価格
            
            ' F12以降(L列～)も、元データのF8以降から2つずらしてコピー
            ' (ExcelのAT列=F46までをカバー)
            On Error Resume Next
            For j = 12 To 46
                rsOut.fields("F" & j).Value = rsIn.fields("F" & (j - 4)).Value
            Next j
            On Error GoTo 0
            
            rsOut.Update
            Debug.Print "Row " & i & ":   └ 【子転写完了】枝番 " & raw_F2
            
        End If
        
        rsIn.MoveNext
    Loop

    rsIn.Close: rsOut.Close
    Debug.Print "=== 全 " & i & " 行の移動・正規化が終了しました ==="
End Sub
