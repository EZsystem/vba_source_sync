Attribute VB_Name = "acc_mod_MappingTemplate"
'-------------------------------------
' Module: acc_mod_MappingTemplate
' 説明   : プロジェクト全体のオブジェクト名管理およびマッピングツール
' 更新日 : 2026/03/26
'-------------------------------------
'Attribute VB_Name = "acc_mod_MappingTemplate"
Option Compare Database
Option Explicit

'===========================================================
' 【RN_ProjectSummary】 オブジェクト定数 最終定義版
'===========================================================
' --- 1. iCube インポート・ワークテーブル ---
Public Const AT_ICUBE As String = "at_Icube"
Public Const AT_ICUBE_HISTORY As String = "at_Icube_累計"
Public Const AT_ICUBE_IMPORT_WORK As String = "at_Temp_Icube_Import"
Public Const AT_ICUBE_COL_SETTING As String = "at_Icube_ColSetting"

' --- 2. バリデーション用マスタ ---
Public Const AT_BUILDING_USE_MAP As String = "at_建物用途正誤表"       ' 旧: tbl_建物用途正誤表
Public Const AT_PRICE_CATEGORY_MAP As String = "at_工事金額区分表"    ' 旧: tbl_工事金額区分表
Public Const AT_CLIENT_DATA As String = "at_顧客データ"                ' 旧: tbl_顧客データ
Public Const AT_BRANCH_WORK_HISTORY As String = "at_支店作業所_累計"   ' 旧: t_支店作業所_累計
Public Const AT_JURISDICTION_MAP As String = "at_管轄作業所_RN部恒久作業所3"      ' 旧: tb_管轄作業所_RN部恒久作業所3
Public Const AT_TEMP_PROJECT_MAP As String = "at_仮基本工事"           ' 旧: tb_仮基本工事
Public Const AT_PROJECT_NAME_CLEAN As String = "at_工事名cle"          ' 旧: tbl_工事名cle
Public Const AT_ERR_SAGYOSHO As String = "at_err作業所"        ' 旧: t_err作業所

' --- 3. 関連マスタ ---
Public Const AT_KIHON_KANKO As String = "at_基本工事_完工"
Public Const AT_KIHON_SAGYO As String = "at_基本工事_作業所"
Public Const AT_KIHON_JUCHU As String = "at_基本工事_受注"
Public Const AT_PROJECT_INFO As String = "at_工事コード情報"
Public Const AT_EDABAN As String = "at_枝番工事"
Public Const AT_LINK_KIHON_NAME As String = "at_基本工事名称_リンク"

' --- 4. ツール管理・クエリ ---
Public Const AT_MAPPING_INFO As String = "at_取込マッピング_Template"
Public Const AQ_SEL_KIHON_NAME As String = "sel_基本工事名称"

' --- 5. 原価管理システム (costManagement_統合分) ---
Public Const AT_GENKA_IMPORT_WORK As String = "at_Temp_原価S_import"
Public Const AT_GENKA_BASIC As String = "at_原価S_基本工事"
Public Const AT_GENKA_BRANCH As String = "at_原価S_枝番工事"
Public Const AT_GENKA_COL_SETTING As String = "at_原価S_ColSetting"
Public Const AT_GENKA_MANUAL_FIX As String = "at_原価S_枝番工事_手動最終補正"

'=================================================
' サブルーチン名 : Generate_MappingTemplate_FromList
' 概要   : マッピングひな型生成のメイン処理
'=================================================
Public Sub Generate_MappingTemplate_FromList()
    Dim db As DAO.Database: Set db = CurrentDb
    
    ' 本テーブル選択
    Dim tableMain As String
    tableMain = Get_TableNameFromList("本テーブルを選択")
    If tableMain = "" Then MsgBox "処理をキャンセルしました", vbExclamation: Exit Sub
    
    ' ひな型生成処理呼び出し (定数 AT_MAPPING_INFO を使用)
    Call Generate_MappingTemplateCore(tableMain, AT_MAPPING_INFO)
End Sub

'=================================================
' 内部関数 : Get_TableNameFromList
'=================================================
Private Function Get_TableNameFromList(promptTitle As String) As String
    Dim db As DAO.Database: Set db = CurrentDb
    Dim td As DAO.TableDef
    Dim tableList As String
    
    For Each td In db.TableDefs
        If Left(td.Name, 4) <> "MSys" Then
            tableList = tableList & td.Name & ";"
        End If
    Next td
    If tableList = "" Then Exit Function
    
    Dim tableArray() As String
    tableArray = Split(Left(tableList, Len(tableList) - 1), ";")
    
    Dim i As Long
    Dim msg As String
    msg = "番号で " & promptTitle & vbCrLf & vbCrLf
    For i = LBound(tableArray) To UBound(tableArray)
        msg = msg & (i + 1) & ". " & tableArray(i) & vbCrLf
    Next i
    
    Dim choice As Variant
    choice = InputBox(msg & vbCrLf & "番号を入力してください : ", promptTitle)
    If IsNumeric(choice) Then
        If choice >= 1 And choice <= UBound(tableArray) + 1 Then
            Get_TableNameFromList = tableArray(choice - 1)
        End If
    End If
End Function

'=================================================
' 内部処理 : Generate_MappingTemplateCore
'=================================================
Private Sub Generate_MappingTemplateCore(tableMain As String, tableTemp As String)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim tdefMain As DAO.TableDef: Set tdefMain = db.TableDefs(tableMain)
    Dim tdefTemp As DAO.TableDef: Set tdefTemp = db.TableDefs(tableTemp)
    
    Dim dictTempFields As Object: Set dictTempFields = CreateObject("Scripting.Dictionary")
    Dim fld As DAO.Field
    For Each fld In tdefTemp.Fields
        dictTempFields(fld.Name) = True
    Next fld
    
    Dim rsMapping As DAO.Recordset
    Set rsMapping = db.OpenRecordset("[" & AT_MAPPING_INFO & "]", dbOpenDynaset)
    
    For Each fld In tdefMain.Fields
        If (fld.Attributes And dbAutoIncrField) Or fld.Type = dbGUID Then GoTo NextField
        
        If Not IsMappingExists(tableMain, fld.Name) Then
            rsMapping.AddNew
            rsMapping!本テーブル名 = tableMain
            rsMapping!本フィールド名 = fld.Name
            rsMapping!仮フィールド名 = IIf(dictTempFields.Exists(fld.Name), fld.Name, "")
            rsMapping!データ型 = GetFieldTypeName(fld.Type)
            rsMapping!取込対象 = True
            rsMapping!デフォルト値 = ""
            rsMapping.Update
        End If
NextField:
    Next fld
    
    rsMapping.Close
    MsgBox "マッピングひな型の生成が完了しました", vbInformation
End Sub

'=================================================
' 内部補助関数
'=================================================
Private Function GetFieldTypeName(fieldType As Integer) As String
    Select Case fieldType
        Case dbBoolean:  GetFieldTypeName = "Yes/No型"
        Case dbByte:     GetFieldTypeName = "バイト型"
        Case dbInteger:  GetFieldTypeName = "整数型"
        Case dbLong:     GetFieldTypeName = "長整数型"
        Case dbSingle:   GetFieldTypeName = "単精度型"
        Case dbDouble:   GetFieldTypeName = "倍精度型"
        Case dbCurrency: GetFieldTypeName = "通貨型"
        Case dbDate:     GetFieldTypeName = "日付/時刻型"
        Case dbText:     GetFieldTypeName = "テキスト型"
        Case dbMemo:     GetFieldTypeName = "メモ型"
        Case Else:       GetFieldTypeName = "未対応型"
    End Select
End Function

Private Function IsMappingExists(tableMain As String, FieldName As String) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT * FROM [" & AT_MAPPING_INFO & "] " & _
        "WHERE [本テーブル名] = '" & Replace(tableMain, "'", "''") & "' " & _
        "AND [本フィールド名] = '" & Replace(FieldName, "'", "''") & "'", _
        dbOpenSnapshot)
    IsMappingExists = Not rs.EOF
    rs.Close
End Function

