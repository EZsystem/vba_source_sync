Attribute VB_Name = "acc_mod_Genka_Main"
'Attribute VB_Name = "acc_mod_Genka_Main"
'----------------------------------------------------------------
' Module: acc_mod_Genka_Main
' 説明   : 原価管理システムにおけるExcelインポート、マッピング転写および検証のメインロジック。
' 更新日 : 2026/03/30
'----------------------------------------------------------------
Option Compare Database
Option Explicit

'===========================================================
' 1. メイン・ワークフロー (Public Procedures)
'===========================================================

'----------------------------------------------------------------
' プロシージャ名 : Run_Genka_Import_Workflow
' 概要          : Excelインポートから本番テーブル転写、補正までの全工程を一括実行
'----------------------------------------------------------------
Public Sub Run_Genka_Import_Workflow()
    Debug.Print "--- 工事原価インポート・ワークフローを開始します ---"
    
    ' 1. Excelファイルのデータ取り込み (A7以降)
    If Not Import_Excel_To_Temp() Then
        MsgBox "Excelインポート処理でエラーが発生したため、中断します。", vbCritical
        Exit Sub
    End If
    
    ' 2. 仮テーブルから本番テーブルへの動的マッピング転写
    Debug.Print "=== 本番テーブルへのデータ転写（DAO/Dynamic Mapping）を開始します ==="
    If Not Transfer_Temp_To_Production() Then
        MsgBox "本番テーブルへの転写処理中にエラーが発生しました。", vbCritical
        Exit Sub
    End If
    
    ' 3. 手動最終補正処理 (枝番コードを軸にした属性修正)
    Debug.Print "=== 手動最終補正処理(Apply_Manual_Final_Correction)を実行します ==="
    Call Apply_Manual_Final_Correction
    
    ' 4. Icube累計との整合性検証
    'Debug.Print "=== Icube累計とのデータ検証を実行します ==="
    'Call Validate_Branch_Against_Icube_Accumulated
    
    MsgBox "すべてのインポート・転写工程が正常に完了しました。", vbInformation
End Sub

'===========================================================
' 2. 工程別内部処理 (Private Procedures)
'===========================================================

'----------------------------------------------------------------
' 内部関数 : Import_Excel_To_Temp
' 概要    : Excelファイルから指定範囲を一時テーブルへインポート
'----------------------------------------------------------------
Private Function Import_Excel_To_Temp() As Boolean
    Import_Excel_To_Temp = False
    Dim db As DAO.Database: Set db = CurrentDb
    Dim xlPath As String
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim lastRow As Long, importRange As String
    
    On Error GoTo Err_Handler
    
    xlPath = "D:\My_code\11_workspaces\RN_kanri_system\genka_system\原価システムimport.xlsm"
    
    ' Excelプロセスを起動して最終行を確認 (xlUp = -4162)
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(xlPath, ReadOnly:=True)
    
    ' オブジェクト名（CodeName）でシートを特定
    Set xlSheet = G_GetSheetByCodeName(xlBook, SH_CODE_IM_GENKA)
    
    ' 見つからない場合はシート名（見出し名）で再試行
    If xlSheet Is Nothing Then
        On Error Resume Next
        Set xlSheet = xlBook.Sheets(SH_NAME_IM_GENKA)
        On Error GoTo Err_Handler
    End If
    
    If xlSheet Is Nothing Then
        MsgBox "対象シートが見つかりません。" & vbCrLf & _
               "オブジェクト名: " & SH_CODE_IM_GENKA & vbCrLf & _
               "シート名: " & SH_NAME_IM_GENKA, vbCritical
        GoTo Exit_Sub
    End If
    
    ' 現在のシート名（見出し名）を取得
    Dim snActual As String
    snActual = xlSheet.Name
    
    lastRow = xlSheet.Cells(xlSheet.rows.count, 1).End(-4162).Row
    xlBook.Close False
    xlApp.Quit
    
    If lastRow < 7 Then
        Debug.Print "A7行以降にデータが存在しません。"
        Exit Function
    End If
    
    ' インポート先ワークテーブルをクリア
    db.Execute "DELETE FROM [" & AT_GENKA_IMPORT_WORK & "]", dbFailOnError
    
    ' A7:AT(最終行)を取得 (動的に取得したシート名を使用)
    importRange = snActual & "!A7:AT" & lastRow
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, _
        AT_GENKA_IMPORT_WORK, xlPath, False, importRange
        
    Debug.Print "  -> インポート完了 (レコード数: " & (lastRow - 6) & ")"
    Import_Excel_To_Temp = True

Exit_Sub:
    Set xlSheet = Nothing: Set xlBook = Nothing: Set xlApp = Nothing
    Exit Function
Err_Handler:
    Debug.Print "Import_Excel_To_Temp Error: " & Err.Description
    Import_Excel_To_Temp = False: Resume Exit_Sub
End Function

'----------------------------------------------------------------
' 内部関数 : Transfer_Temp_To_Production
' 概要    : マッピング設定に基づきDAO経由で本番テーブルへ転送
'----------------------------------------------------------------
Private Function Transfer_Temp_To_Production() As Boolean
    Transfer_Temp_To_Production = False
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsMapCom As DAO.Recordset, rsMapVar As DAO.Recordset, rsIn As DAO.Recordset
    Dim rsKihon As DAO.Recordset, rsEdaban As DAO.Recordset
    Dim dictMap As Object: Set dictMap = CreateObject("Scripting.Dictionary")
    
    ' 本番テーブルのフィールド存在確認用辞書
    Dim validKihon As Object: Set validKihon = CreateObject("Scripting.Dictionary")
    Dim validEdaban As Object: Set validEdaban = CreateObject("Scripting.Dictionary")
    Dim fld As DAO.Field
    
    On Error GoTo Err_Handler
    
    ' 実在チェックリスト作成
    For Each fld In db.TableDefs(AT_GENKA_BASIC).Fields: validKihon(fld.Name) = True: Next
    For Each fld In db.TableDefs(AT_GENKA_BRANCH).Fields: validEdaban(fld.Name) = True: Next
    
    ' マッピング読み込み (共通)
    Set rsMapCom = db.OpenRecordset(AT_GENKA_SETTING_COM, dbOpenSnapshot)
    Call Load_Mapping_To_Dict(rsMapCom, dictMap)
    rsMapCom.Close
    
    ' マッピング読み込み (変数)
    Set rsMapVar = db.OpenRecordset(AT_GENKA_SETTING_VAR, dbOpenSnapshot)
    Call Load_Mapping_To_Dict(rsMapVar, dictMap)
    rsMapVar.Close
    
    ' 転写先クリア
    db.Execute "DELETE FROM [" & AT_GENKA_BASIC & "]", dbFailOnError
    db.Execute "DELETE FROM [" & AT_GENKA_BRANCH & "]", dbFailOnError
    
    Set rsKihon = db.OpenRecordset(AT_GENKA_BASIC, dbOpenDynaset)
    Set rsEdaban = db.OpenRecordset(AT_GENKA_BRANCH, dbOpenDynaset)
    Set rsIn = db.OpenRecordset(AT_GENKA_IMPORT_WORK, dbOpenSnapshot)
    
    ' 作業変数
    Dim v_K1 As String, v_K2 As String, v_K3 As String, v_K4 As String, v_K5 As String, v_K6 As String
    Dim f1_val As String, f2_val As String, f3_val As String
    Dim spacePos As Integer
    Dim countInsK As Long, countUpdK As Long, countInsE As Long
    
    v_K1 = "": v_K2 = "": v_K3 = "": v_K4 = "": v_K5 = "": v_K6 = ""
    
    Do Until rsIn.EOF
        ' --- 行ごとのリセット (持ち越し防止) ---
        v_K5 = "": v_K6 = ""
        
        f1_val = Trim(Nz(rsIn!f1, ""))
        If f1_val = "" Then GoTo NextRow
        
        If f1_val = "1" Then
            ' --- 新規案件開始時のリセット ---
            v_K3 = "": v_K4 = ""
            
            f3_val = Format_Genka_String(Nz(rsIn!f3, ""))
            spacePos = InStr(f3_val, " ")
            v_K1 = IIf(spacePos > 0, Left(f3_val, spacePos - 1), f3_val)
            v_K2 = IIf(spacePos > 0, Mid(f3_val, spacePos + 1), "")
            
            rsKihon.AddNew
            Call Apply_Field_Mapping_Logic(rsKihon, dictMap, validKihon, rsIn, v_K1, v_K2, v_K3, v_K4, v_K5, v_K6)
            rsKihon.Update
            countInsK = countInsK + 1
            
        ElseIf f1_val = "2" Then
            ' 【F1=2】通常の案件情報の更新
            rsKihon.FindFirst "[基本工事コード] = '" & Replace(v_K1, "'", "''") & "'"
            If Not rsKihon.NoMatch Then
                rsKihon.Edit
                Call Apply_Field_Mapping_Logic(rsKihon, dictMap, validKihon, rsIn, v_K1, v_K2, v_K3, v_K4, v_K5, v_K6)
                rsKihon.Update
                countUpdK = countUpdK + 1
            End If
            
        ElseIf f1_val = "3" Then
            ' 【F1=3】K7(経費)の独立更新
            rsKihon.FindFirst "[基本工事コード] = '" & Replace(v_K1, "'", "''") & "'"
            If Not rsKihon.NoMatch Then
                rsKihon.Edit
                ' マッピング外の独立処理: F22を既払高：経費に転写
                rsKihon![既払高：経費] = val(Replace(Replace(Nz(rsIn!f22, "0"), ",", ""), "\", ""))
                rsKihon.Update
                countUpdK = countUpdK + 1
            End If
            
        ElseIf f1_val = "4" Then
            f3_val = Format_Genka_String(Nz(rsIn!f3, ""))
            spacePos = InStr(f3_val, " ")
            v_K3 = IIf(spacePos > 0, Left(f3_val, spacePos - 1), f3_val)
            v_K4 = IIf(spacePos > 0, Mid(f3_val, spacePos + 1), "")
            
        ElseIf val(f1_val) >= 5 Then
            f2_val = Trim(Nz(rsIn!f2, ""))
            If f2_val <> "" And f2_val <> "管理番号" Then
                v_K5 = v_K3 & "-" & f2_val
                v_K6 = Trim(Nz(rsIn!f3, ""))
                
                rsEdaban.AddNew
                Call Apply_Field_Mapping_Logic(rsEdaban, dictMap, validEdaban, rsIn, v_K1, v_K2, v_K3, v_K4, v_K5, v_K6)
                rsEdaban.Update
                countInsE = countInsE + 1
            End If
        End If
NextRow:
        rsIn.MoveNext
    Loop
    
    Debug.Print "  基本追加: " & countInsK & " / 基本更新: " & countUpdK & " / 枝番追加: " & countInsE
    Transfer_Temp_To_Production = True

Exit_Sub:
    On Error Resume Next
    rsIn.Close: rsKihon.Close: rsEdaban.Close: db.Close
    Exit Function
Err_Handler:
    Debug.Print "Transfer_Temp_To_Production Error: " & Err.Description
    Resume Exit_Sub
End Function

'----------------------------------------------------------------
' 内部処理 : Apply_Field_Mapping_Logic
'----------------------------------------------------------------
Private Sub Apply_Field_Mapping_Logic(rsDest As DAO.Recordset, dictMap As Object, validFields As Object, rsData As DAO.Recordset, _
                                     v_K1 As String, v_K2 As String, v_K3 As String, _
                                     v_K4 As String, v_K5 As String, v_K6 As String)
    Dim key As Variant, param() As String, destField As String, destType As String
    Dim rawVal As Variant
    
    ' dictMap(K1-K6, およびFフィールド)を順番に処理
    For Each key In dictMap.Keys
        ' K7はF1=3ブロックで別途個別処理するため、マッピングからは除外
        If CStr(key) <> "K7" Then
            param = Split(dictMap(key), "|")
            destField = param(0): destType = param(1)
            
            If validFields.Exists(destField) Then
                rawVal = Null
                If Left(CStr(key), 1) = "K" Then
                    Select Case CStr(key)
                        Case "K1": rawVal = v_K1: Case "K2": rawVal = v_K2
                        Case "K3": rawVal = v_K3: Case "K4": rawVal = v_K4
                        Case "K5": rawVal = v_K5: Case "K6": rawVal = v_K6
                    End Select
                Else
                    rawVal = rsData.Fields(CStr(key)).Value
                End If
                
                Call Apply_OneField_To_Dest(rsDest, destField, destType, rawVal)
            End If
        End If
    Next key
End Sub

' 1フィールドを実際に代入する補助プロシージャ
Private Sub Apply_OneField_To_Dest(ByRef rsDest As DAO.Recordset, ByVal destField As String, ByVal destType As String, ByVal rawVal As Variant)
    Dim strVal As String
    If Not IsNull(rawVal) Then
        strVal = Trim(CStr(rawVal))
        If strVal <> "" Then
            If Right(strVal, 1) = "%" Then
                strVal = Replace(strVal, "%", "")
                If IsNumeric(strVal) Then rsDest.Fields(destField).Value = CDbl(strVal) / 100
            Else
                If InStr(destType, "通貨") > 0 Or InStr(destType, "倍精度") > 0 Or InStr(destType, "数値") > 0 Then
                    strVal = Replace(Replace(strVal, ",", ""), "\", "")
                    If IsNumeric(strVal) Then rsDest.Fields(destField).Value = strVal
                Else
                    rsDest.Fields(destField).Value = strVal
                End If
            End If
        End If
    End If
End Sub

'----------------------------------------------------------------
' 内部補助 : Load_Mapping_To_Dict
'----------------------------------------------------------------
Private Sub Load_Mapping_To_Dict(rs As DAO.Recordset, ByRef dict As Object)
    Dim fld As DAO.Field, colMoto As String, colSaki As String, colType As String
    colMoto = "": colSaki = "": colType = ""
    
    ' 列名の動的特定
    For Each fld In rs.Fields
        If InStr(fld.Name, "元") > 0 Or InStr(fld.Name, "変数") > 0 Then colMoto = fld.Name
        If InStr(fld.Name, "先") > 0 Or InStr(fld.Name, "タイトル") > 0 Then colSaki = fld.Name
        If InStr(fld.Name, "型") > 0 Then colType = fld.Name
    Next fld
    
    Do Until rs.EOF
        If Trim(Nz(rs.Fields(colMoto).Value, "")) <> "" Then
            dict(Trim(rs.Fields(colMoto).Value)) = Trim(Nz(rs.Fields(colSaki).Value, "")) & "|" & Trim(Nz(rs.Fields(colType).Value, ""))
        End If
        rs.MoveNext
    Loop
End Sub

'----------------------------------------------------------------
' 内部補助 : Format_Genka_String
'----------------------------------------------------------------
Private Function Format_Genka_String(ByVal src As String) As String
    src = Replace(Replace(src, vbTab, " "), "　", " ")
    Do While InStr(src, "  ") > 0: src = Replace(src, "  ", " "): Loop
    Format_Genka_String = Trim(src)
End Function

'===========================================================
' 3. 検証・補正処理 (Validation & Correction)
'===========================================================

'----------------------------------------------------------------
' プロシージャ名 : Apply_Manual_Final_Correction
' 概要          : 枝番工事コードを検索軸として、属性情報をマスタ正解値で差し替える
'----------------------------------------------------------------
Public Sub Apply_Manual_Final_Correction()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim strSQL As String
    Const TEMP_CORR As String = "at_Temp_Genka_Correction_Work"

    On Error GoTo Err_Sub
    
    On Error Resume Next: db.Execute "DROP TABLE [" & TEMP_CORR & "]", dbFailOnError: On Error GoTo Err_Sub

    ' 1. 退避
    strSQL = "SELECT B.* INTO [" & TEMP_CORR & "] " & _
             "FROM [" & AT_GENKA_BRANCH & "] AS B " & _
             "INNER JOIN [" & AT_GENKA_MANUAL_FIX & "] AS M ON B.[枝番工事コード] = M.[枝番コード];"
    db.Execute strSQL, dbFailOnError

    ' 2. 属性修正
    strSQL = "UPDATE [" & TEMP_CORR & "] AS T " & _
             "INNER JOIN [" & AT_GENKA_MANUAL_FIX & "] AS M ON T.[枝番工事コード] = M.[枝番コード] " & _
             "SET T.[工事コード] = M.[工事コード], T.[管理番号] = M.[管理番号], T.[追加工事名称] = M.[追加工事名称];"
    db.Execute strSQL, dbFailOnError

    ' 3. 本番からの削除と復元
    db.Execute "DELETE FROM [" & AT_GENKA_BRANCH & "] WHERE [枝番工事コード] IN (SELECT [枝番コード] FROM [" & AT_GENKA_MANUAL_FIX & "]);", dbFailOnError
    db.Execute "INSERT INTO [" & AT_GENKA_BRANCH & "] SELECT * FROM [" & TEMP_CORR & "];", dbFailOnError

    db.Execute "DROP TABLE [" & TEMP_CORR & "]", dbFailOnError
    Debug.Print "  -> 手動最終補正完了"
    Exit Sub
Err_Sub:
    Debug.Print "Apply_Manual_Final_Correction Error: " & Err.Description
End Sub

'----------------------------------------------------------------
' プロシージャ名 : Validate_Branch_Against_Icube_Accumulated
' 概要          : 枝番工事+工事価格のペアが Icube累計に存在するか検証し、エラーを記録
'----------------------------------------------------------------
Public Sub Validate_Branch_Against_Icube_Accumulated()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim sql As String
    On Error GoTo Err_Handler
    
    sql = "UPDATE [" & AT_GENKA_BRANCH & "] AS E " & _
          "LEFT JOIN [" & AT_ICUBE_HISTORY & "] AS I " & _
          "ON (E.[枝番工事コード] & E.[工事価格]) = (I.[枝番工事コード] & I.[工事価格]) " & _
          "SET E.[枝番工事コードerr] = 'Icubeと枝番工事コード不一致' " & _
          "WHERE I.[枝番工事コード] Is Null;"
    db.Execute sql, dbFailOnError
    Debug.Print "  -> Icube累計との整合性検証完了"
    Exit Sub
Err_Handler:
    Debug.Print "Validation Error: " & Err.Description
End Sub






