Attribute VB_Name = "acc_mod_MappingTemplate"
Option Explicit

'----------------------------------------------------------------
' Module: acc_mod_MappingTemplate
' 説明   : 職員勤怠・兼務率システム用オブジェクト定義およびマッピングロジック
' 更新日 : 2026/04/01
'----------------------------------------------------------------

'===========================================================
' 1. テーブル・クエリ名称定義 (Table/Query Names)
'===========================================================
Public Const AT_KENMU_TEMP As String = "at_kenmuTemp" ' 暫定取り込み用
Public Const AT_KENMU_MAIN As String = "at_kenmu"     ' 本番データ用
Public Const AT_KENMU_HISTORY As String = "at_kenmu_累計" ' 累計データ用
Public Const AT_TEMP_PROJECT_MAP As String = "_at_仮基本工事" ' 仮基本工事マッピング用
Public Const AT_ORG_MAP As String = "_at_MapOrgRN3" ' 組織マッピング用
Public Const AT_STAFF_MAIN As String = "at_社員情報"
Public Const AT_SYSTEM_REG As String = "_at_SystemRegistry"

'===========================================================
' 2. Excel シート・テーブル名称定義 (Excel Objects)
'===========================================================
Public Const SH_NAME_KENMU As String = "職員兼務率" ' 取得対象シート名
Public Const LO_NAME_KENMU As String = "xt_kenmu"   ' 取得対象リストオブジェクト名

'===========================================================
' 3. データ整形・クレンジングロジック (Mapping Logic)
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
    
    ' 重複キーを避けるため、既存のキーがあれば上書き（最新）とするか無視するか
    ' ここでは辞書を作成して返す
    Set rs = db.OpenRecordset("SELECT [作業所_略称], [施工管轄組織名] FROM [" & AT_ORG_MAP & "]", dbOpenSnapshot)
    
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
