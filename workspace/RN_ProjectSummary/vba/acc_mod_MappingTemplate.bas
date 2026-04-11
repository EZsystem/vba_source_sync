Attribute VB_Name = "acc_mod_MappingTemplate"
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
Public Const AT_ORDER_FORECAST_KEISU As String = "at_受注額予測計数"
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

' --- 小口工事予測 (at_Work_01-05) ---
Public Const AT_WORK_01_ACTUALS_3P As String = "at_Work_01_実績推移_3期分"
Public Const AT_WORK_02_ORDER_3P_AVE As String = "at_Work_02_受注_3期平均"
Public Const AT_WORK_03_COMP_RATIO As String = "at_Work_03_完工_推移割合"
Public Const AT_WORK_04_ORDER_FCST As String = "at_Work_04_受注_今期予測"
Public Const AT_WORK_05_COMP_FCST As String = "at_Work_05_完工_今期予測"

' --- 最終集計結果 ---
Public Const AT_WORK_FINAL_AGGREGATION As String = "at_Work_給与経費集計_結果"

' --- Excel Sheet CodeNames & Sheet Names (内部オブジェクト名と見出し名) ---
Public Const SH_CODE_EX_MASTER      As String = "sh_Ex_Master"     ' 経費M (オブジェクト名)
Public Const SH_NAME_EX_MASTER      As String = "経費M"            ' 経費M (シート名)

Public Const SH_CODE_EX_ICUBE       As String = "sh_Ex_Icube"      ' IcubeData (オブジェクト名)
Public Const SH_NAME_EX_ICUBE       As String = "IcubeData"        ' IcubeData (シート名)

Public Const SH_CODE_EX_GENKA       As String = "sh_Ex_Genka"      ' 原価Data (エクスポート用オブジェクト名)
Public Const SH_NAME_EX_GENKA       As String = "原価Data"         ' 原価Data (エクスポート用シート名)

Public Const SH_CODE_IM_GENKA       As String = "CostMng1"         ' 原価システムimport (インポート用オブジェクト名)
Public Const SH_NAME_IM_GENKA       As String = "原価S直データ"    ' 原価システムimport (インポート用シート名)

Public Const SH_CODE_EX_EMP         As String = "sh_EX_Emp"        ' 兼務率2 (オブジェクト名)
Public Const SH_CODE_EX_FORECAST    As String = "sh_EX_Forecast"   ' ac_受注完工予測 (オブジェクト名)
Public Const SH_NAME_EX_FORECAST    As String = "ac_受注完工予測"  ' ac_受注完工予測 (シート名)

' --- 職員兼務率用 ---
Public Const SH_NAME_KENMU          As String = "職員兼務率"        ' 職員兼務率 (シート名)
Public Const LO_NAME_KENMU          As String = "xt_kenmu"         ' 職員兼務率 (テーブル名)

' --- ツール管理・クエリ ---
Public Const AT_MAPPING_INFO As String = "at_取込マッピング_Template"
Public Const AQ_SEL_KIHON_NAME As String = "sel_基本工事名称"
Public Const AQ_SMALL_PROJECT_TRANS As String = "q_小口受注完工推移_3期分"
Public Const AQ_ORDER_FORECAST_WA As String = "q_受注完工予測_加重平均集計"

' --- 原価管理システム (Cost Management) ---
Public Const AT_GENKA_IMPORT_WORK As String = "at_Temp_原価S_import"
Public Const AT_GENKA_BASIC As String = "at_原価S_基本工事"
Public Const AT_GENKA_BRANCH As String = "at_原価S_枝番工事"
Public Const AT_GENKA_SETTING_COM As String = "at_原価S_ColSettingCom"
Public Const AT_GENKA_SETTING_VAR As String = "at_原価S_ColSettingVar"
Public Const AT_GENKA_MANUAL_FIX As String = "at_原価S_枝番工事_手動最終補正"

' --- エクスポート管理 ---
Public Const AT_EXPORT_CONFIG As String = "_at_ExportConfig"

' --- 職員兼務率・各種マスタ ---
Public Const AT_KENMU_TEMP As String = "at_kenmuTemp" ' 暫定取り込み用
Public Const AT_KENMU_MAIN As String = "at_kenmu"     ' 本番データ用
Public Const AT_KENMU_HISTORY As String = "at_kenmu_累計" ' 累計データ用
Public Const AT_STAFF_MAIN As String = "at_社員情報"
Public Const AT_SYSTEM_REG As String = "_at_SystemRegistry"

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

'===========================================================
' 3. データ整形・クレンジングロジック (兼務率等 Mapping Logic)
'===========================================================

'--------------------------------------------
' 関数名 : Cleanse_Percent_Smart
' 概要   : 兼務率（50%, 50, 0.5）を 0.5 という小数に統一する
'--------------------------------------------
Public Function Cleanse_Percent_Smart(ByVal val As Variant) As Double
    Dim sRaw As String: sRaw = Trim(Nz(val, ""))
    
    If sRaw = "" Or sRaw = "0" Then
        Cleanse_Percent_Smart = 0
        Exit Function
    End If
    
    ' 数値として扱えるかチェック
    Dim dVal As Double
    If IsNumeric(Replace(sRaw, "%", "")) Then
        dVal = CDbl(Replace(sRaw, "%", ""))
        
        ' パターン判定
        If InStr(sRaw, "%") > 0 Then
            ' "50%" -> 0.5
            Cleanse_Percent_Smart = dVal / 100
        ElseIf dVal > 1 Then
            ' "50" -> 0.5 (1より大きい場合は整数表記のパーセントとみなす)
            Cleanse_Percent_Smart = dVal / 100
        Else
            ' "0.5" -> 0.5 (1以下の場合は既に小数表記とみなす)
            Cleanse_Percent_Smart = dVal
        End If
    Else
        ' 数字ですらない場合は 0 を返す
        Cleanse_Percent_Smart = 0
    End If
End Function

'--------------------------------------------
' 関数名 : Cleanse_Date_Smart
' 概要   : 日付文字列を「月の初日」の日付型に変換する
'--------------------------------------------
Public Function Cleanse_Date_Smart(ByVal val As Variant) As Variant
    Dim sRaw As String: sRaw = Trim(Nz(val, ""))
    If sRaw = "" Then Exit Function
    
    If IsDate(sRaw) Then
        ' 月の初日に補正
        Cleanse_Date_Smart = DateSerial(Year(CDate(sRaw)), Month(CDate(sRaw)), 1)
        Exit Function
    End If
    
    ' 和暦等の特殊形式への簡易対応
    Dim sConv As String
    sConv = Replace(Replace(sRaw, "年", "/"), "月", "")
    If Right(sConv, 1) <> "/" Then sConv = sConv & "/1"
    
    If IsDate(sConv) Then
        Cleanse_Date_Smart = CDate(sConv)
    Else
        Cleanse_Date_Smart = Null
    End If
End Function

'--------------------------------------------
' 関数名 : Get_FiscalTerm
' 概要   : 日付から「期（年度）」を計算する（4月開始）
'          例：2026/04 -> 14
'--------------------------------------------
Public Function Get_FiscalTerm(ByVal dt As Variant) As String
    If Not IsDate(dt) Then Get_FiscalTerm = "": Exit Function
    
    Dim d As Date: d = CDate(dt)
    Dim y As Long: y = Year(d)
    Dim m As Long: m = Month(d)
    
    ' 4月?12月なら Year - 2012, 1月?3月なら Year - 2013
    If m >= 4 Then
        Get_FiscalTerm = CStr(y - 2012) & "期"
    Else
        Get_FiscalTerm = CStr(y - 2013) & "期"
    End If
End Function

'--------------------------------------------
' 関数名 : Get_Quarter
' 概要   : 日付から「Q（四半期）」を計算する（4月開始）
'          例：4-6月 -> 1Q, 1-3月 -> 4Q
'--------------------------------------------
Public Function Get_Quarter(ByVal dt As Variant) As String
    If Not IsDate(dt) Then Get_Quarter = "": Exit Function
    
    Dim m As Long: m = Month(CDate(dt))
    
    Select Case m
        Case 4 To 6:   Get_Quarter = "1Q"
        Case 7 To 9:   Get_Quarter = "2Q"
        Case 10 To 12: Get_Quarter = "3Q"
        Case 1 To 3:   Get_Quarter = "4Q"
        Case Else:     Get_Quarter = ""
    End Select
End Function

'--------------------------------------------
' 関数名 : Get_TempProject_Map
' 概要   : 仮基本工事マッピングデータを一括取得する
'          戻り値：(0 to n, 0 to 1) の 2次元配列
'                  (n, 0): 仮基本工事名称 (ワイルドカード込)
'                  (n, 1): 仮基本工事コード
'--------------------------------------------
Public Function Get_TempProject_Map() As Variant
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim arr() As String
    Dim i As Long: i = 0
    
    Set rs = db.OpenRecordset("SELECT [仮基本工事名称], [仮基本工事コード] FROM [" & AT_TEMP_PROJECT_MAP & "]", dbOpenSnapshot)
    
    If Not rs.EOF Then
        rs.MoveLast: rs.MoveFirst
        ReDim arr(rs.recordCount - 1, 1)
        
        Do Until rs.EOF
            ' 名称内の 「？？（全角）」 を 「??（半角）」 に置換して検索用に最適化
            arr(i, 0) = Replace(Nz(rs![仮基本工事名称], ""), "？", "?")
            arr(i, 1) = Nz(rs![仮基本工事コード], "")
            i = i + 1
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    
    Get_TempProject_Map = arr
End Function

'--------------------------------------------
' 関数名 : Get_Org_Dict
' 概要   : 作業所略称（半角）をキー、施工管轄組織名を値とする辞書を返す
'--------------------------------------------
Public Function Get_Org_Dict() As Object
    Dim db   As DAO.Database: Set db = CurrentDb
    Dim rs   As DAO.Recordset
    Dim dict As Object
    Dim sKey As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' AT_JURISDICTION_MAPに変更 (at_管轄作業所_RN部恒久作業所3)
    Set rs = db.OpenRecordset("SELECT [作業所_略称], [施工管轄組織名] FROM [" & AT_JURISDICTION_MAP & "]", dbOpenSnapshot)
    
    Do Until rs.EOF
        sKey = StrConv(Nz(rs![作業所_略称], ""), vbNarrow) ' 半角化してキーにする
        If sKey <> "" Then
            If Not dict.Exists(sKey) Then
                dict.Add sKey, Nz(rs![施工管轄組織名], "")
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    Set Get_Org_Dict = dict
End Function
