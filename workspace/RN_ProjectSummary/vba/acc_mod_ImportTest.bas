Attribute VB_Name = "acc_mod_ImportTest"
'Attribute VB_Name = "acc_mod_ImportTest"
Option Compare Database
Option Explicit

' 転写関連のテーブル名を一元管理(定数宣言)
Private Const TBL_RAW_IMPORT As String = "at_Test_Raw_Import"
Private Const TBL_SETTING_COM As String = "at_原価S_ColSettingCom"
Private Const TBL_SETTING_VAR As String = "at_原価S_ColSettingVar"
Private Const TBL_KIHON As String = "at_原価S_基本工事"
Private Const TBL_EDABAN As String = "at_原価S_枝番工事"

'============================================
' プロシージャ名 : Execute_All_Genka_Process
' 概要          : インポートから本テーブルへの転写(仮テーブルからの動的SQL)
'               までの一連の処理を順次実行する。
'============================================
Public Sub Execute_All_Genka_Process()
    Debug.Print "全体の処理を開始します..."
    
    If Not Import_Genka_Raw_Test() Then
        MsgBox "インポート処理でエラーが発生したため、後続のデータ加工を中止します。", vbCritical
        Exit Sub
    End If
    
    ' 古い正規化ロジック(Test2, Test3)は使用せず、一気に転写処理を実行します
    Debug.Print "=== 新ロジックによる本テーブルへのデータ転写（UPDATE）を開始します ==="
    
    If Not Update_Genka_MainTables_FromTemp() Then
        MsgBox "本テーブルへの転写処理中にエラーが発生しました。", vbCritical
        Exit Sub
    End If
    
    ' ---- [追加] モジュール外部の手動補正処理を呼び出し ----
    Debug.Print "=== 手動最終補正処理(Apply_Manual_Final_Correction)を実行します ==="
    Call Apply_Manual_Final_Correction
    
    ' ---- [追加] Icube累計とのデータ検証処理を呼び出し ----
    'Debug.Print "=== Icubeデータとの枝番工事の検証を実行します ==="
    'Call Validate_Edaban_Against_Icube
    
    MsgBox "すべての転写処理が正常に完了しました。", vbInformation
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
    lastRow = xlSheet.Cells(xlSheet.Rows.count, 1).End(-4162).Row
    
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
' プロシージャ名 : Update_Genka_MainTables_FromTemp
' 概要          : 共通設定テーブル(at_原価S_ColSettingCom)と
'               変数設定テーブル(at_原価S_ColSettingVar)のマッピングに基づき、
'               本テーブルに該当フィールドが実在するかを確認した上で、
'               仮テーブルのデータを転写(UPDATE)する
'============================================
Public Function Update_Genka_MainTables_FromTemp() As Boolean
    Update_Genka_MainTables_FromTemp = False
    Dim db As DAO.Database
    Dim rsMapCom As DAO.Recordset
    Dim rsMapVar As DAO.Recordset
    Dim rsIn As DAO.Recordset
    Dim tdfKihon As DAO.TableDef
    Dim tdfEdaban As DAO.TableDef
    Dim fld As DAO.Field
    Dim df_fldKari As String, df_destFld As String, df_destType As String
    
    Dim dictMap As Object
    Set dictMap = CreateObject("Scripting.Dictionary")
    
    Dim fieldNamesKihon As Object
    Dim fieldNamesEdaban As Object
    Set fieldNamesKihon = CreateObject("Scripting.Dictionary")
    Set fieldNamesEdaban = CreateObject("Scripting.Dictionary")
    
    Dim v_K1 As String, v_K2 As String, v_K3 As String
    Dim v_K4 As String, v_K5 As String, v_K6 As String
    Dim f1_val As String, f2_val As String, f3_val As String
    Dim spacePos As Integer
    Dim rsKihon As DAO.Recordset
    Dim rsEdaban As DAO.Recordset
    
    On Error GoTo Err_Handler
    Set db = CurrentDb
    
    ' 2. 本テーブルのフィールド構造を「実在チェックリスト」として辞書に格納
    Set tdfKihon = db.TableDefs(TBL_KIHON)
    For Each fld In tdfKihon.Fields
        fieldNamesKihon(fld.Name) = True
    Next fld
    
    Set tdfEdaban = db.TableDefs(TBL_EDABAN)
    For Each fld In tdfEdaban.Fields
        fieldNamesEdaban(fld.Name) = True
    Next fld
    
    ' --- フォールバック用のフィールド取得関数(インライン) ---
    Dim fldMap As DAO.Field
    Dim k_moto As String, k_saki As String, k_kata As String
    
    ' 2. マッピングの読み込み (共通テーブル: at_原価S_ColSettingCom)
    Debug.Print "=== マッピング(共通)を読み込みます ==="
    Set rsMapCom = db.OpenRecordset("SELECT * FROM [" & TBL_SETTING_COM & "]", dbOpenSnapshot)
    
    ' 列名を動的に探す(タイポ対策)
    Dim colComMoto As String, colComSaki As String, colComType As String
    For Each fldMap In rsMapCom.Fields
        If InStr(fldMap.Name, "元") > 0 Or InStr(fldMap.Name, "変数") > 0 Or InStr(fldMap.Name, "仮テーブル") > 0 Then colComMoto = fldMap.Name
        If InStr(fldMap.Name, "先") > 0 Or InStr(fldMap.Name, "タイトル") > 0 Then colComSaki = fldMap.Name
        If InStr(fldMap.Name, "型") > 0 Then colComType = fldMap.Name
    Next fldMap
    
    Do Until rsMapCom.EOF
        df_fldKari = Trim(Nz(rsMapCom.Fields(colComMoto).Value, ""))
        If df_fldKari <> "" Then
            df_destFld = Trim(Nz(rsMapCom.Fields(colComSaki).Value, ""))
            df_destType = Trim(Nz(rsMapCom.Fields(colComType).Value, ""))
            If df_destFld <> "" Then
                dictMap(df_fldKari) = df_destFld & "|" & df_destType
            End If
        End If
        rsMapCom.MoveNext
    Loop
    rsMapCom.Close
    
    ' 3. マッピングの読み込み (変数テーブル: at_原価S_ColSettingVar)
    Debug.Print "=== マッピング(変数)を読み込みます ==="
    Set rsMapVar = db.OpenRecordset("SELECT * FROM [" & TBL_SETTING_VAR & "]", dbOpenSnapshot)
    
    ' 列名を動的に探す(タイポ対策)
    Dim colVarMoto As String, colVarSaki As String, colVarType As String
    For Each fldMap In rsMapVar.Fields
        If InStr(fldMap.Name, "元") > 0 Or InStr(fldMap.Name, "変数") > 0 Or InStr(fldMap.Name, "仮テーブル") > 0 Then colVarMoto = fldMap.Name
        If InStr(fldMap.Name, "先") > 0 Or InStr(fldMap.Name, "タイトル") > 0 Then colVarSaki = fldMap.Name
        If InStr(fldMap.Name, "型") > 0 Then colVarType = fldMap.Name
    Next fldMap
    
    Do Until rsMapVar.EOF
        df_fldKari = Trim(Nz(rsMapVar.Fields(colVarMoto).Value, ""))
        If df_fldKari <> "" Then
            df_destFld = Trim(Nz(rsMapVar.Fields(colVarSaki).Value, ""))
            df_destType = Trim(Nz(rsMapVar.Fields(colVarType).Value, ""))
            If df_destFld <> "" Then
                dictMap(df_fldKari) = df_destFld & "|" & df_destType
            End If
        End If
        rsMapVar.MoveNext
    Loop
    rsMapVar.Close
    
    ' 5. 【追加】転写前に対象の本番テーブルをクリアする
    Debug.Print "=== 転写先のテーブルデータをクリアします ==="
    db.Execute "DELETE FROM [" & TBL_KIHON & "]", dbFailOnError
    db.Execute "DELETE FROM [" & TBL_EDABAN & "]", dbFailOnError
    
    Set rsKihon = db.OpenRecordset(TBL_KIHON, dbOpenDynaset)
    Set rsEdaban = db.OpenRecordset(TBL_EDABAN, dbOpenDynaset)
    
    Dim insertKihonCount As Long, updateKihonCount As Long, insertEdabanCount As Long
    insertKihonCount = 0: updateKihonCount = 0: insertEdabanCount = 0
    
    ' 6. 仮テーブルのループ処理
    Set rsIn = db.OpenRecordset("SELECT * FROM [" & TBL_RAW_IMPORT & "]", dbOpenSnapshot)
    Debug.Print "=== マッピングと実在チェックに基づく新しい安全なデータ転写処理(DAO方式)を開始します ==="
    
    Do Until rsIn.EOF
        f1_val = Trim(Nz(rsIn!f1, ""))
        If f1_val = "" Then
            ' ブランクやヘッダー空行はスキップして次へ
            GoTo NextRow
        End If
        
        If f1_val = "1" Then
            f3_val = Trim(Nz(rsIn!f3, ""))
            f3_val = Replace(f3_val, vbTab, " ")
            f3_val = Replace(f3_val, "　", " ")
            Do While InStr(f3_val, "  ") > 0
                f3_val = Replace(f3_val, "  ", " ")
            Loop
            
            spacePos = InStr(f3_val, " ")
            If spacePos > 0 Then
                v_K1 = Trim(Left(f3_val, spacePos - 1))
                v_K2 = Trim(Mid(f3_val, spacePos + 1))
            Else
                v_K1 = f3_val
                v_K2 = ""
            End If
            
            ' F1=1のとき、基本工事を新規追加(INSERT)
            rsKihon.AddNew
            Call ApplyMappingToRS(rsKihon, dictMap, fieldNamesKihon, rsIn, v_K1, v_K2, v_K3, v_K4, v_K5, v_K6)
            rsKihon.Update
            insertKihonCount = insertKihonCount + 1
            
        ElseIf f1_val = "2" Or f1_val = "3" Then
            ' F1=2, 3のとき、さきほど追加した基本工事に対して追記(UPDATE)
            rsKihon.FindFirst "[基本工事コード] = '" & Replace(v_K1, "'", "''") & "'"
            If Not rsKihon.NoMatch Then
                rsKihon.Edit
                Call ApplyMappingToRS(rsKihon, dictMap, fieldNamesKihon, rsIn, v_K1, v_K2, v_K3, v_K4, v_K5, v_K6)
                rsKihon.Update
                updateKihonCount = updateKihonCount + 1
            End If
            
        ElseIf f1_val = "4" Then
            f3_val = Trim(Nz(rsIn!f3, ""))
            f3_val = Replace(f3_val, vbTab, " ")
            f3_val = Replace(f3_val, "　", " ")
            Do While InStr(f3_val, "  ") > 0
                f3_val = Replace(f3_val, "  ", " ")
            Loop
            
            spacePos = InStr(f3_val, " ")
            If spacePos > 0 Then
                v_K3 = Trim(Left(f3_val, spacePos - 1))
                v_K4 = Trim(Mid(f3_val, spacePos + 1))
            Else
                v_K3 = f3_val
                v_K4 = ""
            End If
            
        ElseIf val(f1_val) >= 5 Then
            f2_val = Trim(Nz(rsIn!f2, ""))
            
            ' F2(管理番号) が空文字、またはヘッダーの「管理番号」の場合は転写の対象外とする
            If f2_val <> "" And f2_val <> "管理番号" Then
                v_K5 = v_K3 & "‐" & f2_val
                v_K6 = Trim(Nz(rsIn!f3, ""))
                
                ' F1>=5のとき、枝番工事を新規追加(INSERT)
                rsEdaban.AddNew
                Call ApplyMappingToRS(rsEdaban, dictMap, fieldNamesEdaban, rsIn, v_K1, v_K2, v_K3, v_K4, v_K5, v_K6)
                rsEdaban.Update
                insertEdabanCount = insertEdabanCount + 1
            End If
        End If
        
NextRow:
        rsIn.MoveNext
    Loop
    
    Debug.Print "=== UPDATE処理が正常に完了しました ==="
    Debug.Print "  基本工事 新規追加 : " & insertKihonCount & " 件"
    Debug.Print "  基本工事 追記更新 : " & updateKihonCount & " 件"
    Debug.Print "  枝番工事 新規追加 : " & insertEdabanCount & " 件"
    
    Update_Genka_MainTables_FromTemp = True

Exit_Sub:
    On Error Resume Next
    Set dictMap = Nothing
    Set fieldNamesKihon = Nothing
    Set fieldNamesEdaban = Nothing
    If Not rsMapCom Is Nothing Then rsMapCom.Close
    If Not rsMapVar Is Nothing Then rsMapVar.Close
    If Not rsIn Is Nothing Then rsIn.Close
    If Not rsKihon Is Nothing Then rsKihon.Close
    If Not rsEdaban Is Nothing Then rsEdaban.Close
    Set rsMapCom = Nothing
    Set rsMapVar = Nothing
    Set rsIn = Nothing
    Set rsKihon = Nothing
    Set rsEdaban = Nothing
    Set tdfKihon = Nothing
    Set tdfEdaban = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Debug.Print "エラー (" & Err.Number & "): " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "転写UPDATE処理エラー"
    Update_Genka_MainTables_FromTemp = False
    Resume Exit_Sub
End Function

'============================================
' プロシージャ名 : ApplyMappingToRS
' 概要          : マッピングと物理列確認に基づき、Recordsetに値をセットする(DAO方式)
'============================================
Private Sub ApplyMappingToRS(rsDest As DAO.Recordset, dictMap As Object, validFields As Object, rsData As DAO.Recordset, _
                             v_K1 As String, v_K2 As String, v_K3 As String, _
                             v_K4 As String, v_K5 As String, v_K6 As String)
    Dim key As Variant, mapInfo As String
    Dim param() As String
    Dim destField As String, destType As String
    Dim rawVal As Variant, strVal As String
    
    For Each key In dictMap.Keys
        mapInfo = CStr(dictMap(key))
        param = Split(mapInfo, "|")
        destField = param(0)
        destType = param(1)
        
        ' 転写先テーブルにこのフィールド名が実在するかをチェック
        If validFields.Exists(destField) Then
            rawVal = Null
            
            ' 取得元の値を決定
            If Left(CStr(key), 1) = "K" Then
                Select Case CStr(key)
                    Case "K1": rawVal = v_K1
                    Case "K2": rawVal = v_K2
                    Case "K3": rawVal = v_K3
                    Case "K4": rawVal = v_K4
                    Case "K5": rawVal = v_K5
                    Case "K6": rawVal = v_K6
                End Select
            Else
                rawVal = rsData.Fields(CStr(key)).Value
            End If
            
            If Not IsNull(rawVal) Then
                strVal = Trim(CStr(rawVal))
                If strVal <> "" Then
                    ' %があれば数値変換(0.xx)
                    If Right(strVal, 1) = "%" Then
                        strVal = Replace(strVal, "%", "")
                        If IsNumeric(strVal) Then
                            rsDest.Fields(destField).Value = CDbl(strVal) / 100
                        End If
                    Else
                        ' %以外はDAOに代入を任せる。ただし数字のカンマや円記号などは削除する
                        If InStr(destType, "通貨") > 0 Or InStr(destType, "倍精度") > 0 Or InStr(destType, "数値") > 0 Then
                             strVal = Replace(strVal, ",", "")
                             strVal = Replace(strVal, "\", "")
                             strVal = Replace(strVal, "\", "")
                             rsDest.Fields(destField).Value = strVal
                        Else
                             rsDest.Fields(destField).Value = strVal
                        End If
                    End If
                End If
            End If
        End If
    Next key
End Sub

'============================================
' プロシージャ名 : Debug_Print_KihonName_Test
' 概要          : F1=1の行のF3, F4などの値をイミディエイトに出力し、
'               基本工事名がどこに格納されているかを調査するテストコード
'============================================
Public Sub Debug_Print_KihonName_Test()
    Dim db As DAO.Database
    Dim rsIn As DAO.Recordset
    Dim f3_val As String, f4_val As String
    Dim spacePos As Integer
    Dim v_K1 As String, v_K2 As String
    Dim count As Integer
    
    Set db = CurrentDb
    ' F1='1' の行を取得
    Set rsIn = db.OpenRecordset("SELECT * FROM [at_Test_Raw_Import] WHERE Nz(F1, '') = '1'")
    
    Debug.Print "=== 調査開始：F1=1 の基本工事名取得テスト ==="
    
    Do Until rsIn.EOF Or count >= 5
        f3_val = Trim(Nz(rsIn!f3, ""))
        f4_val = Trim(Nz(rsIn!F4, ""))
        
        Debug.Print "【確認 " & (count + 1) & "件目 】"
        Debug.Print "  元のF3の内容: [" & f3_val & "]"
        Debug.Print "  元のF4の内容: [" & f4_val & "]"
        
        ' ロジックと同じ分割処理を実行
        f3_val = Replace(f3_val, vbTab, " ")
        f3_val = Replace(f3_val, "　", " ")
        Do While InStr(f3_val, "  ") > 0
            f3_val = Replace(f3_val, "  ", " ")
        Loop
        
        spacePos = InStr(f3_val, " ")
        If spacePos > 0 Then
            v_K1 = Trim(Left(f3_val, spacePos - 1))
            v_K2 = Trim(Mid(f3_val, spacePos + 1))
        Else
            v_K1 = f3_val
            v_K2 = ""
        End If
        
        Debug.Print "  --> 構築された 基本工事コード(K1): [" & v_K1 & "]"
        Debug.Print "  --> 構築された 基本工事名　(K2): [" & v_K2 & "]"
        Debug.Print "---------------------------------------------------"
        
        count = count + 1
        rsIn.MoveNext
    Loop
    
    rsIn.Close
    Set rsIn = Nothing
    Set db = Nothing
    Debug.Print "=== 調査完了 ==="
End Sub

'============================================
' プロシージャ名 : Debug_Print_MappingMatch_Test
' 概要          : マッピングテーブル(Com, Var)から読み取った値と、
'               実際の転写先(基本・枝番)にその列が実在するかを
'               イミディエイトに出力して検証する小テストコード
'============================================
Public Sub Debug_Print_MappingMatch_Test()
    Dim db As DAO.Database
    Dim rsMap As DAO.Recordset
    Dim tdfKihon As DAO.TableDef
    Dim tdfEdaban As DAO.TableDef
    Dim fld As DAO.Field
    Dim fldMap As DAO.Field
    
    Dim fieldNamesKihon As Object
    Dim fieldNamesEdaban As Object
    Set fieldNamesKihon = CreateObject("Scripting.Dictionary")
    Set fieldNamesEdaban = CreateObject("Scripting.Dictionary")
    
    Set db = CurrentDb
    
    Debug.Print "========== 設定テーブル 突き合わせテスト =========="
    
    ' 1. 本テーブルのフィールドを確認
    Set tdfKihon = db.TableDefs("at_原価S_基本工事")
    For Each fld In tdfKihon.Fields
        fieldNamesKihon(fld.Name) = True
    Next fld
    
    Set tdfEdaban = db.TableDefs("at_原価S_枝番工事")
    For Each fld In tdfEdaban.Fields
        fieldNamesEdaban(fld.Name) = True
    Next fld
    
    Debug.Print "【1. at_原価S_ColSettingComの突き合わせ】"
    On Error Resume Next
    Set rsMap = db.OpenRecordset("SELECT * FROM [at_原価S_ColSettingCom]", dbOpenSnapshot)
    If Err.Number <> 0 Then
        Debug.Print "  -> エラー: テーブルが見つかりません"
    Else
        Call Check_Mapping_Match(rsMap, fieldNamesKihon, fieldNamesEdaban)
        rsMap.Close
    End If
    On Error GoTo 0
    
    Debug.Print ""
    Debug.Print "【2. at_原価S_ColSettingVarの突き合わせ】"
    On Error Resume Next
    Set rsMap = db.OpenRecordset("SELECT * FROM [at_原価S_ColSettingVar]", dbOpenSnapshot)
    If Err.Number <> 0 Then
        Debug.Print "  -> エラー: テーブルが見つかりません"
    Else
        Call Check_Mapping_Match(rsMap, fieldNamesKihon, fieldNamesEdaban)
        rsMap.Close
    End If
    On Error GoTo 0
    
    Set tdfKihon = Nothing
    Set tdfEdaban = Nothing
    Set db = Nothing
    Debug.Print "========== テスト終了 =========="
End Sub

Private Sub Check_Mapping_Match(rsMap As DAO.Recordset, kihonDict As Object, edabanDict As Object)
    Dim k_moto As String, k_saki As String, k_kata As String
    Dim fldMap As DAO.Field
    Dim df_Moto As String, df_Saki As String, df_Type As String
    Dim isKihonExist As String, isEdabanExist As String
    
    ' 列名の自動特定
    For Each fldMap In rsMap.Fields
        If InStr(fldMap.Name, "元") > 0 Or InStr(fldMap.Name, "変数") > 0 Or InStr(fldMap.Name, "仮テーブル") > 0 Then k_moto = fldMap.Name
        If InStr(fldMap.Name, "先") > 0 Or InStr(fldMap.Name, "タイトル") > 0 Then k_saki = fldMap.Name
        If InStr(fldMap.Name, "型") > 0 Then k_kata = fldMap.Name
    Next fldMap
    
    If k_moto = "" Or k_saki = "" Or k_kata = "" Then
        Debug.Print "  -> エラー：このテーブルから「元」「先」「型」という文字を含む列名を見つけられませんでした。"
        Exit Sub
    End If
    
    Debug.Print "読取対象列: [" & k_moto & "] -> [" & k_saki & "] (型: " & k_kata & ")"
    
    Do Until rsMap.EOF
        df_Moto = Trim(Nz(rsMap.Fields(k_moto).Value, ""))
        If df_Moto <> "" Then
            df_Saki = Trim(Nz(rsMap.Fields(k_saki).Value, ""))
            df_Type = Trim(Nz(rsMap.Fields(k_kata).Value, ""))
            
            ' 存在チェック
            If kihonDict.Exists(df_Saki) Then
                isKihonExist = "〇(実在する)"
            Else
                isKihonExist = "×(無視される)"
            End If
            
            If edabanDict.Exists(df_Saki) Then
                isEdabanExist = "〇(実在する)"
            Else
                isEdabanExist = "×(無視される)"
            End If
            
            Debug.Print "  [" & df_Moto & "]->[" & df_Saki & "] | 基本工事: " & isKihonExist & " | 枝番工事: " & isEdabanExist
        End If
        rsMap.MoveNext
    Loop
End Sub

'============================================
' プロシージャ名 : Validate_Edaban_Against_Icube
' 概要          : 枝番工事テーブルの「枝番工事コード＋追加工事名称」が
'               at_Icube_累計 に存在するかを検証する。存在しない場合は
'               枝番工事コードerr フィールドにエラー文言を書き込む
'============================================
Public Sub Validate_Edaban_Against_Icube()
    Dim db As DAO.Database
    Dim sql As String
    
    On Error GoTo Err_Handler
    Set db = CurrentDb
    
    ' LEFT JOIN を用いて、at_Icube_累計 に一致するレコードが無い at_原価S_枝番工事 の行を抽出して UPDATE
    sql = "UPDATE [" & TBL_EDABAN & "] AS E " & _
          "LEFT JOIN [at_Icube_累計] AS I " & _
          "ON (E.[枝番工事コード] & E.[追加工事名称]) = " & _
          "(I.[枝番工事コード] & I.[追加工事名称]) " & _
          "SET E.[枝番工事コードerr] = 'Icubeと枝番工事コード不一致' " & _
          "WHERE I.[枝番工事コード] Is Null;"
          
    db.Execute sql, dbFailOnError
    
    Debug.Print "  -> Icube検証(枝番工事不一致チェック)完了"
    
Exit_Sub:
    Set db = Nothing
    Exit Sub
    
Err_Handler:
    Debug.Print "Icube検証エラー (" & Err.Number & "): " & Err.Description
    Resume Exit_Sub
End Sub


