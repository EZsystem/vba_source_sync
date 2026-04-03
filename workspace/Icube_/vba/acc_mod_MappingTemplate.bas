Attribute VB_Name = "acc_mod_MappingTemplate"
'-------------------------------------
' Module: accmod_MappingTemplate
' 説明   : マッピングひな型生成処理
' 作成日 : 2025/05/23
' 更新日 : -
'-------------------------------------
Option Compare Database
Option Explicit

'=================================================
' 定数名 : MAPPING_TABLE_NAME
' 説明   : マッピングテーブル名を定義（変更時ここだけ修正）
'=================================================
Private Const MAPPING_TABLE_NAME As String = "tb_本テーブル取込マッピング"

'=================================================
' サブルーチン名 : Generate_MappingTemplate_FromList
' 説明   : 一覧から本テーブルを選択し
'        : マッピングひな型生成処理を呼び出す
' 引数   : なし
' 戻り値 : なし
'=================================================
Public Sub Generate_MappingTemplate_FromList()
    ' --- 1. DAO.Database 取得 ---
    Dim Db As DAO.Database
    Set Db = CurrentDb
    
    ' --- 2. 本テーブル選択 ---
    Dim tableMain As String
    tableMain = Get_TableNameFromList("本テーブルを選択")
    If tableMain = "" Then MsgBox "処理をキャンセルしました", vbExclamation: Exit Sub
    
    ' --- 3. 仮テーブル選択 ---
    Dim tableTemp As String
    tableTemp = MAPPING_TABLE_NAME
    
    ' --- 4. ひな型生成処理呼び出し ---
    Call Generate_MappingTemplateCore(tableMain, tableTemp)
End Sub ' ← Generate_MappingTemplate_FromList 終了

'=================================================
' 関数名 : Get_TableNameFromList
' 説明   : テーブル名一覧から選択したテーブル名を取得
' 引数   : promptTitle（String）プロンプトタイトル
' 戻り値 : String 選択したテーブル名
'=================================================
Private Function Get_TableNameFromList(promptTitle As String) As String
    ' --- 1. DAO.Database 取得 ---
    Dim Db As DAO.Database
    Set Db = CurrentDb
    
    ' --- 2. テーブル名一覧作成 ---
    Dim td As DAO.TableDef
    Dim tableList As String
    For Each td In Db.TableDefs
        If Left(td.Name, 4) <> "MSys" Then
            tableList = tableList & td.Name & ";"
        End If
    Next td
    If tableList = "" Then Exit Function
    
    Dim tableArray() As String
    tableArray = Split(Left(tableList, Len(tableList) - 1), ";")
    
    ' --- 3. プロンプト作成 ---
    Dim i As Long
    Dim msg As String
    msg = "番号で " & promptTitle & vbCrLf & vbCrLf
    For i = LBound(tableArray) To UBound(tableArray)
        msg = msg & (i + 1) & ". " & tableArray(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "番号を入力してください : "
    
    ' --- 4. 入力値検証 ---
    Dim choice As Variant
    choice = InputBox(msg, promptTitle)
    If IsNumeric(choice) Then
        If choice >= 1 And choice <= UBound(tableArray) + 1 Then
            Get_TableNameFromList = tableArray(choice - 1)
        End If
    End If
End Function ' ← Get_TableNameFromList 終了

'=================================================
' サブルーチン名 : Generate_MappingTemplateCore
' 説明   : 指定本テーブルと仮テーブルからフィールド情報を取得し
'        : マッピングテーブルのひな型を生成
' 引数   : tableMain（String）本テーブル名
'        : tableTemp（String）仮テーブル名
' 戻り値 : なし
'=================================================
Private Sub Generate_MappingTemplateCore(tableMain As String, tableTemp As String)
    ' --- 1. DAO.Database 取得 ---
    Dim Db As DAO.Database
    Set Db = CurrentDb
    
    ' --- 2. テーブル定義取得 ---
    Dim tdefMain As DAO.TableDef
    Dim tdefTemp As DAO.TableDef
    Set tdefMain = Db.TableDefs(tableMain)
    Set tdefTemp = Db.TableDefs(tableTemp)
    
    ' --- 3. 仮テーブルフィールド辞書作成 ---
    Dim dictTempFields As Object
    Set dictTempFields = CreateObject("Scripting.Dictionary")
    Dim fld As DAO.Field
    For Each fld In tdefTemp.fields
        dictTempFields(fld.Name) = True
    Next fld
    
    ' --- 4. マッピングテーブルレコードセット取得 ---
    Dim rsMapping As DAO.Recordset
    Set rsMapping = Db.OpenRecordset(MAPPING_TABLE_NAME, dbOpenDynaset)
    
    ' --- 5. 本テーブルフィールドループ ---
    Dim fldName As String
    Dim fieldTypeStr As String
    Dim existsInTemp As Boolean
    For Each fld In tdefMain.fields
        fldName = fld.Name
        
        ' 自動増分フィールドと GUID 型はスキップ
        If (fld.Attributes And dbAutoIncrField) Or fld.Type = dbGUID Then GoTo NextField
        
        fieldTypeStr = GetFieldTypeName(fld.Type)
        existsInTemp = dictTempFields.Exists(fldName)
        
        If Not IsMappingExists(tableMain, fldName) Then
            rsMapping.AddNew
            rsMapping!本テーブル名 = tableMain
            rsMapping!本フィールド名 = fldName
            rsMapping!仮フィールド名 = IIf(existsInTemp, fldName, "")
            rsMapping!データ型 = fieldTypeStr
            rsMapping!取込対象 = True
            rsMapping!デフォルト値 = ""
            rsMapping.Update
        End If
NextField:
    Next fld
    
    ' --- 6. クリーンアップ ---
    rsMapping.Close
    
    ' --- 7. 完了通知 ---
    MsgBox "マッピングひな型の生成が完了しました", vbInformation
End Sub ' ← Generate_MappingTemplateCore 終了

'=================================================
' 関数名 : GetFieldTypeName
' 説明   : フィールド型定数から日本語の型名を返す
' 引数   : fieldType（Integer）DAO.Field.Type 定数
' 戻り値 : String 型名称
'=================================================
Private Function GetFieldTypeName(fieldType As Integer) As String
    Select Case fieldType
        Case dbBoolean:   GetFieldTypeName = "Yes/No型"
        Case dbByte:      GetFieldTypeName = "バイト型"
        Case dbInteger:   GetFieldTypeName = "整数型"
        Case dbLong:      GetFieldTypeName = "長整数型"
        Case dbSingle:    GetFieldTypeName = "単精度型"
        Case dbDouble:    GetFieldTypeName = "倍精度型"
        Case dbCurrency:  GetFieldTypeName = "通貨型"
        Case dbDate:      GetFieldTypeName = "日付/時刻型"
        Case dbText:      GetFieldTypeName = "テキスト型"
        Case dbMemo:      GetFieldTypeName = "メモ型"
        Case dbGUID:      GetFieldTypeName = "GUID型"
        Case Else:        GetFieldTypeName = "未対応型"
    End Select
End Function ' ← GetFieldTypeName 終了

'=================================================
' 関数名 : IsMappingExists
' 説明   : 指定本テーブルとフィールド名でマッピングが既存か判定
' 引数   : tableMain（String）本テーブル名
'        : fieldName（String）フィールド名
' 戻り値 : Boolean 既存なら True
'=================================================
Private Function IsMappingExists(tableMain As String, FieldName As String) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT * FROM [" & MAPPING_TABLE_NAME & "] " & _
        "WHERE [本テーブル名] = '" & Replace(tableMain, "'", "''") & "' " & _
        "AND [本フィールド名] = '" & Replace(FieldName, "'", "''") & "'", _
        dbOpenSnapshot)
    IsMappingExists = Not rs.EOF
    rs.Close
End Function ' ← IsMappingExists 終了


