Attribute VB_Name = "acc_mod_MappingTemplate"
'Attribute VB_Name = "acc_mod_MappingTemplate"
'----------------------------------------------------------------
' Module: acc_mod_Genka_MappingManager (Original: acc_mod_MappingTemplate)
' 説明   : 原価管理システム用オブジェクト名の定義、およびマッピング管理ユーティリティ。
' 更新日 : 2026/03/30
'----------------------------------------------------------------
Option Compare Database
Option Explicit

'===========================================================
' 1. オブジェクト定数定義 (Table/Query Names)
'===========================================================

' --- iCube インポート・ワークテーブル ---
Public Const AT_ICUBE As String = "at_Icube"
Public Const AT_ICUBE_HISTORY As String = "at_Icube_累計"
Public Const AT_ICUBE_IMPORT_WORK As String = "at_Temp_Icube_Import"
Public Const AT_ICUBE_COL_SETTING As String = "at_Icube_ColSetting"

' --- バリデーション用マスタ ---
Public Const AT_BUILDING_USE_MAP As String = "at_建物用途正誤表"
Public Const AT_PRICE_CATEGORY_MAP As String = "at_工事金額区分表"
Public Const AT_CLIENT_DATA As String = "at_顧客データ"
Public Const AT_BRANCH_WORK_HISTORY As String = "at_支店作業所_累計"
Public Const AT_JURISDICTION_MAP As String = "at_管轄作業所_RN部恒久作業所3"
Public Const AT_TEMP_PROJECT_MAP As String = "at_仮基本工事"
Public Const AT_PROJECT_NAME_CLEAN As String = "at_工事名cle"
Public Const AT_ERR_SAGYOSHO As String = "at_err作業所"

' --- 関連マスタ ---
Public Const AT_KIHON_KANKO As String = "at_基本工事_完工"
Public Const AT_KIHON_SAGYO As String = "at_基本工事_作業所"
Public Const AT_KIHON_JUCHU As String = "at_基本工事_受注"
Public Const AT_PROJECT_INFO As String = "at_工事コード情報"
Public Const AT_EDABAN As String = "at_枝番工事"
Public Const AT_LINK_KIHON_NAME As String = "at_基本工事名称_リンク"

' --- Excel Sheet CodeNames (不変のオブジェクト名) ---
Public Const SH_CODE_VALUE_ENTRY   As String = "sh_ValueEntry"    ' 14期_計算入力
Public Const SH_CODE_CALC_WORK     As String = "sh_CalcWork"      ' 中間計算ワーク
Public Const SH_CODE_REPORT_OUT    As String = "sh_ReportOutput"  ' 管理表出力
Public Const SH_CODE_EX_MASTER      As String = "sh_Ex_Master"     ' M_経費/組織 (旧: xt)
Public Const SH_CODE_EX_ICUBE       As String = "sh_Ex_Icube"      ' Icubeデータ
Public Const SH_CODE_EX_GENKA       As String = "sh_Ex_Genka"      ' 原価データ
Public Const SH_CODE_EX_EMP         As String = "sh_EX_Emp"        ' 兼務率_給与人員 (旧: 兼務率2)

' --- ツール管理・クエリ ---
Public Const AT_MAPPING_INFO As String = "at_取込マッピング_Template"
Public Const AQ_SEL_KIHON_NAME As String = "sel_基本工事名称"

' --- 原価管理システム (Cost Management) ---
Public Const AT_GENKA_IMPORT_WORK As String = "at_Temp_原価S_import"
Public Const AT_GENKA_BASIC As String = "at_原価S_基本工事"
Public Const AT_GENKA_BRANCH As String = "at_原価S_枝番工事"
Public Const AT_GENKA_SETTING_COM As String = "at_原価S_ColSettingCom"
Public Const AT_GENKA_SETTING_VAR As String = "at_原価S_ColSettingVar"
Public Const AT_GENKA_MANUAL_FIX As String = "at_原価S_枝番工事_手動最終補正"

' --- エクスポート管理 ---
Public Const AT_EXPORT_CONFIG As String = "_at_ExportConfig"

'===========================================================
' 2. マッピングひな型生成ユーティリティ (Development Tools)
'===========================================================

'----------------------------------------------------------------
' ユーティリティ名 : Run_Create_Mapping_Template_Tables
' 概要            : 指定のテーブルを基に、マッピング用のひな型データを生成
'----------------------------------------------------------------
Public Sub Run_Create_Mapping_Template_Tables()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim tableMain As String
    
    ' テーブル選択の疑似UI (InputBox)
    tableMain = Get_TableName_From_User("マッピング対象の本番テーブルを選択してください")
    If tableMain = "" Then Exit Sub
    
    Call Build_Mapping_Template_Core(tableMain)
End Sub

Private Function Get_TableName_From_User(title As String) As String
    Dim td As DAO.TableDef, msg As String, i As Long, tableArray() As String
    For Each td In CurrentDb.TableDefs
        If Left(td.Name, 4) <> "MSys" Then msg = msg & (i + 1) & ". " & td.Name & vbCrLf: i = i + 1
    Next
    If i = 0 Then Exit Function
    
    ReDim tableArray(i - 1): i = 0
    For Each td In CurrentDb.TableDefs
        If Left(td.Name, 4) <> "MSys" Then tableArray(i) = td.Name: i = i + 1
    Next
    
    Dim choice As Variant: choice = InputBox(msg & vbCrLf & "番号を入力:", title)
    If IsNumeric(choice) Then
        If choice >= 1 And choice <= UBound(tableArray) + 1 Then Get_TableName_From_User = tableArray(choice - 1)
    End If
End Function

Private Sub Build_Mapping_Template_Core(tableMain As String)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim tdefMain As DAO.TableDef: Set tdefMain = db.TableDefs(tableMain)
    Dim rsMapping As DAO.Recordset: Set rsMapping = db.OpenRecordset("[" & AT_MAPPING_INFO & "]", dbOpenDynaset)
    Dim fld As DAO.Field
    
    For Each fld In tdefMain.Fields
        If (fld.Attributes And dbAutoIncrField) Or fld.Type = dbGUID Then GoTo NextF
        
        ' 存在チェック (簡易)
        Dim checkRS As DAO.Recordset
        Set checkRS = db.OpenRecordset("SELECT * FROM [" & AT_MAPPING_INFO & "] WHERE [本テーブル名]='" & tableMain & "' AND [本フィールド名]='" & fld.Name & "'")
        If checkRS.EOF Then
            rsMapping.AddNew
            rsMapping!本テーブル名 = tableMain
            rsMapping!本フィールド名 = fld.Name
            rsMapping!データ型 = Get_Field_Type_Literal(fld.Type)
            rsMapping!取込対象 = True
            rsMapping.Update
        End If
        checkRS.Close
NextF:
    Next fld
    rsMapping.Close
    MsgBox "マッピングひな型の生成が完了しました (" & tableMain & ")", vbInformation
End Sub

Private Function Get_Field_Type_Literal(t As Integer) As String
    Select Case t
        Case dbBoolean: Get_Field_Type_Literal = "Yes/No型": Case dbLong: Get_Field_Type_Literal = "長整数型"
        Case dbDouble: Get_Field_Type_Literal = "倍精度型": Case dbCurrency: Get_Field_Type_Literal = "通貨型"
        Case dbDate: Get_Field_Type_Literal = "日付/時刻型": Case dbText: Get_Field_Type_Literal = "テキスト型"
        Case Else: Get_Field_Type_Literal = "その他"
    End Select
End Function


