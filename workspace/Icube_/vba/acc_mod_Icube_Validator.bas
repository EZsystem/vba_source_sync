Attribute VB_Name = "acc_mod_Icube_Validator"
'Attribute VB_Name = "acc_mod_Icube_Validator"
'===================================================================================================
' モジュール名 : acc_mod_Icube_Validator
' 概要         : Icube_テーブル固有のバリデーション・判定・補完ロジック（Phase 1-8）
' 依存関係     : acc_clsDataCleaner, com_clsErrorUtility
' 最終更新日   : 2026/03/26
'===================================================================================================

Option Compare Database
Option Explicit

' 会計年度計算用の定数
Private Const BASE_YEAR As Integer = 2012

'---------------------------------------------------------------------------------------------------
' 1. 司令塔(Main)から直接呼び出される Public プロシージャ
'---------------------------------------------------------------------------------------------------

'===========================================================
' プロシージャ名 : Process_BasicValidation_And_Split
' 概要           : Phase 1-2: 基本バリデーションと名称分割
'===========================================================
Public Sub Process_BasicValidation_And_Split(ByRef Cleaner As acc_clsDataCleaner, ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    
    ' Phase 1: 基本クレンジングと補完
    Call Process_Judge_OneTimeProject
    Call Process_Copy_Empty_ProjectInfo
    Call Process_Merge_BranchCode
    Call Process_DateConversion_Smart(Cleaner)
    Call Process_Update_Jurisdiction(ErrorLog)
    
    ' Phase 2: 分割とマッピング
    Call Process_Split_ProjectNames
    Call Process_Calculate_PeriodFromName
    Call Process_Transfer_TempProjectCode
    Call Process_Map_OrderFieldsToIcube
    
    Exit Sub

Err_Handler:
    ErrorLog.Notify_Smart_Popup "Validator Phase 1-2 Error: " & Err.Description, "Error", vbCritical
End Sub

'===========================================================
' プロシージャ名 : Process_Category_And_Price
' 概要           : Phase 3-4: 用途補正と金額区分
'===========================================================
Public Sub Process_Category_And_Price(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim DbObj As DAO.Database: Set DbObj = CurrentDb
    Dim RsMain As DAO.Recordset
    Dim RsMap As DAO.Recordset
    Dim RawText As String
    Dim ProjectPrice As Currency

    ' Phase 3: 用途補正ロジック
    Set RsMain = DbObj.OpenRecordset("Icube_", dbOpenDynaset)
    Set RsMap = DbObj.OpenRecordset("tbl_建物用途正誤表", dbOpenSnapshot)
    
    Do While Not RsMain.EOF
        RawText = Trim(Nz(RsMain!用途大区分, ""))
        RsMap.MoveFirst
        Do While Not RsMap.EOF
            If RawText = Trim(Nz(RsMap!誤_用途大区分, "")) Then
                RsMain.Edit
                RsMain!s用途大区分 = Trim(Nz(RsMap!正_用途大区分, ""))
                RsMain!s用途大区分名 = Trim(Nz(RsMap!正_用途大区分名, ""))
                RsMain.Update
                Exit Do
            End If
            RsMap.MoveNext
        Loop
        RsMain.MoveNext
    Loop
    RsMain.Close

    ' Phase 4: 金額区分割当
    Set RsMain = DbObj.OpenRecordset("Icube_", dbOpenDynaset)
    Set RsMap = DbObj.OpenRecordset("tbl_工事金額区分表", dbOpenSnapshot)
    
    Do While Not RsMain.EOF
        ProjectPrice = CCur(Nz(RsMain!工事価格, 0))
        RsMap.MoveFirst
        Do While Not RsMap.EOF
            If ProjectPrice >= CCur(Nz(RsMap!最小金額, 0)) And ProjectPrice <= CCur(Nz(RsMap!最大金額, 0)) Then
                RsMain.Edit
                RsMain!工事金額区分コード = RsMap!工事金額区分コード
                RsMain!工事金額区分名 = RsMap!工事金額区分名
                RsMain!工事金額マイナス判定 = RsMap!工事金額マイナス判定
                RsMain.Update
                Exit Do
            End If
            RsMap.MoveNext
        Loop
        RsMain.MoveNext
    Loop
    RsMain.Close
    
    Set RsMain = Nothing
    Set RsMap = Nothing
    Exit Sub

Err_Handler:
    ErrorLog.Notify_Smart_Popup "Validator Phase 3-4 Error: " & Err.Description
End Sub

'===========================================================
' プロシージャ名 : Process_Transcribe_ProjectInfo
' 概要           : Phase 5-6: 基本工事情報転写
'===========================================================
Public Sub Process_Transcribe_ProjectInfo(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim DbObj As DAO.Database: Set DbObj = CurrentDb
    Dim RsTarget As DAO.Recordset
    Dim ProjectCode As String

    ' Phase 5: デフォルト転写
    DbObj.Execute "UPDATE Icube_ SET s基本工事コード = 基本工事コード, s基本工事名称 = 基本工事名称 WHERE 基本工事コード IS NOT NULL", dbFailOnError

    ' Phase 6: スキップリスト判定転写
    Set RsTarget = DbObj.OpenRecordset("SELECT No, s基本工事コード, 工事コード, 工事帳票名, s基本工事名称 FROM Icube_", dbOpenDynaset)
    Do While Not RsTarget.EOF
        ProjectCode = Trim(UCase(Nz(RsTarget!s基本工事コード, "")))
        If Not ProjectCode Like "KT*" Then
            RsTarget.Edit
            RsTarget!s基本工事コード = RsTarget!工事コード
            RsTarget!s基本工事名称 = RsTarget!工事帳票名
            RsTarget.Update
        End If
        RsTarget.MoveNext
    Loop
    RsTarget.Close
    Exit Sub

Err_Handler:
    ErrorLog.Notify_Smart_Popup "Validator Phase 5-6 Error: " & Err.Description
End Sub

'===========================================================
' プロシージャ名 : Process_Final_Cleansing
' 概要           : Phase 7-8: 名称整形と顧客データ転写
'===========================================================
Public Sub Process_Final_Cleansing(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim RsT As DAO.Recordset
    Dim rsS As DAO.Recordset
    Dim OriginalName As String
    Dim CleanedResult As String
    Dim ClientCode As String

    ' Phase 7: 追加工事名称整形
    Set RsT = CurrentDb.OpenRecordset("SELECT No, 発注者コード, 追加工事名称, 追加工事名称_cle FROM Icube_", dbOpenDynaset)
    Do While Not RsT.EOF
        OriginalName = Nz(RsT!追加工事名称, "")
        CleanedResult = GetCleanedName_FromMaster(Nz(RsT!発注者コード, ""), OriginalName)
        
        CleanedResult = Replace(Replace(Replace(CleanedResult, " ", ""), "　", ""), vbTab, "")
        Do While InStr(CleanedResult, "【") > 0 And InStr(CleanedResult, "】") > InStr(CleanedResult, "【")
            CleanedResult = Left(CleanedResult, InStr(CleanedResult, "【") - 1) & Mid(CleanedResult, InStr(CleanedResult, "】") + 1)
        Loop
        CleanedResult = StrConv(CleanedResult, vbWide)
        CleanedResult = Replace(CleanedResult, "��", "(株)")
        
        RsT.Edit
        RsT!追加工事名称_cle = CleanedResult
        RsT.Update
        RsT.MoveNext
    Loop
    RsT.Close

    ' Phase 8: 顧客データ転写
    Set rsS = CurrentDb.OpenRecordset("tbl_顧客データ", dbOpenSnapshot)
    Do While Not rsS.EOF
        ClientCode = Nz(rsS!顧客コード, "")
        If ClientCode <> "" Then
            Set RsT = CurrentDb.OpenRecordset("SELECT 発注者名_tbl FROM Icube_ WHERE 発注者コード = '" & ClientCode & "' AND (発注者名_tbl IS NULL OR 発注者名_tbl = '')", dbOpenDynaset)
            Do While Not RsT.EOF
                RsT.Edit
                RsT!発注者名_tbl = rsS!会社名
                RsT.Update
                RsT.MoveNext
            Loop
            RsT.Close
        End If
        rsS.MoveNext
    Loop
    rsS.Close
    Exit Sub

Err_Handler:
    ErrorLog.Notify_Smart_Popup "Validator Phase 7-8 Error: " & Err.Description
End Sub

'---------------------------------------------------------------------------------------------------
' 2. 内部補助プロシージャ (Private)
'---------------------------------------------------------------------------------------------------

Private Sub Process_Judge_OneTimeProject()
    Dim rs As DAO.Recordset: Set rs = CurrentDb.OpenRecordset("Icube_", dbOpenDynaset)
    Dim Conds As Variant: Conds = Array("１２諸工事", "１３諸工事", "１Ｑ", "２Ｑ", "３Ｑ", "４Ｑ")
    Dim i As Integer
    Dim IsSmall As Boolean
    Do While Not rs.EOF
        IsSmall = False
        For i = 0 To UBound(Conds)
            If InStr(1, Nz(rs!基本工事名称, ""), Conds(i), vbTextCompare) > 0 Then
                IsSmall = True
                Exit For
            End If
        Next i
        rs.Edit
        rs!一件工事判定 = IIf(IsSmall, "小口工事", "一件工事")
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub Process_Copy_Empty_ProjectInfo()
    CurrentDb.Execute "UPDATE Icube_ SET 基本工事コード = 工事コード WHERE 基本工事コード IS NULL OR 基本工事コード = 'N/A'", dbFailOnError
    CurrentDb.Execute "UPDATE Icube_ SET 基本工事名称 = 工事帳票名 WHERE 基本工事名称 IS NULL OR 基本工事名称 = '' OR 基本工事名称 = 'N/A'", dbFailOnError
End Sub

Private Sub Process_Merge_BranchCode()
    CurrentDb.Execute "UPDATE Icube_ SET 枝番工事コード = Nz(工事コード,'') & '-' & Nz(工事枝番,'')", dbFailOnError
End Sub

Private Sub Process_DateConversion_Smart(ByRef Cleaner As acc_clsDataCleaner)
    Dim rs As DAO.Recordset
    Dim Flds As Variant: Flds = Array("[データ年月（受注計上年月）]", "[完成年月日（枝番単位）]")
    Dim Prfx As Variant: Prfx = Array("受注", "完工")
    Dim i As Integer, TargetFld As String, TargetDate As Date
    
    For i = 0 To UBound(Flds)
        TargetFld = IIf(Prfx(i) = "受注", "受注計上日_日付型", Prfx(i) & "日_日付型")
        ' 一件工事判定と基本工事名称を追加で取得（小口工事の名称優先ロジックのため）
        Set rs = CurrentDb.OpenRecordset("SELECT No, " & Flds(i) & ", " & Prfx(i) & "年度, " & Prfx(i) & "期, " & Prfx(i) & "Q, " & Prfx(i) & "月, " & TargetFld & ", 一件工事判定, 基本工事名称 FROM Icube_", dbOpenDynaset)
        
        Do While Not rs.EOF
            ' acc_clsDataCleaner を使用して安全に日付を取得
            TargetDate = Cleaner.CleanDate(rs.fields(1).Value)
            If VBA.Year(TargetDate) > 1900 Then
                rs.Edit
                rs.fields(Prfx(i) & "年度").Value = GetFiscalYear(TargetDate)
                
                ' --- 完工期かつ小口工事の場合のみ名称から算出 (RN_ProjectSummary基準) ---
                Dim isSmallHandled As Boolean: isSmallHandled = False
                If Prfx(i) = "完工" And rs!一件工事判定 = "小口工事" Then
                    Dim projName As String: projName = Nz(rs!基本工事名称, "")
                    Dim posYear As Long: posYear = InStr(projName, "年度")
                    If posYear >= 3 Then
                        Dim yearVal As Integer
                        yearVal = Val(StrConv(Mid(projName, posYear - 2, 2), vbNarrow))
                        rs.fields(Prfx(i) & "期").Value = yearVal - 12
                        isSmallHandled = True
                    End If
                End If
                
                ' 通常ロジック (一件工事、または受注期の場合、あるいは名称から取得できなかった場合)
                If Not isSmallHandled Then
                    ' 修正: +1 を除去し、RN_ProjectSummary と統一 (2025年度 = 13期)
                    rs.fields(Prfx(i) & "期").Value = GetFiscalYear(TargetDate) - BASE_YEAR
                End If
                
                rs.fields(Prfx(i) & "Q").Value = GetFiscalQuarter(TargetDate)
                rs.fields(Prfx(i) & "月").Value = VBA.Month(TargetDate)
                rs.fields(TargetFld).Value = TargetDate
                rs.Update
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next i
End Sub

Private Sub Process_Update_Jurisdiction(ByRef ErrorLog As com_clsErrorUtility)
    Dim Db As DAO.Database: Set Db = CurrentDb
    Dim rsD As DAO.Recordset, RsT As DAO.Recordset, RsE As DAO.Recordset
    Dim Dict As Object: Set Dict = CreateObject("Scripting.Dictionary")
    
    Set rsD = Db.OpenRecordset("SELECT 組織コード, 施工管轄組織コード FROM t_支店作業所_累計", dbOpenSnapshot)
    Do While Not rsD.EOF
        If Not IsNull(rsD!組織コード) Then
            Dict(Trim(CStr(rsD!組織コード))) = rsD!施工管轄組織コード
        End If
        rsD.MoveNext
    Loop
    rsD.Close
    
    Set RsT = Db.OpenRecordset("Icube_", dbOpenDynaset)
    Set RsE = Db.OpenRecordset("t_err作業所", dbOpenDynaset)
    Do While Not RsT.EOF
        Dim oC As String: oC = Trim(Nz(RsT!施工担当組織コード, ""))
        If Dict.Exists(oC) Then
            RsT.Edit
            RsT!施工管轄組織コード = Dict(oC)
            RsT.Update
        Else
            RsE.AddNew
            RsE!追加工事名称 = RsT!追加工事名称
            RsE!施工担当組織コード = RsT!施工担当組織コード
            RsE!施工担当組織名 = RsT!施工担当組織名
            RsE.Update
        End If
        RsT.MoveNext
    Loop
    RsT.Close: RsE.Close
    
    Set rsD = Db.OpenRecordset("SELECT 施工管轄組織コード, 施工管轄組織名 FROM tb_管轄作業所_RN部恒久作業所3", dbOpenSnapshot)
    Dict.RemoveAll
    Do While Not rsD.EOF
        If Not IsNull(rsD!施工管轄組織コード) Then
            Dict(Trim(CStr(rsD!施工管轄組織コード))) = rsD!施工管轄組織名
        End If
        rsD.MoveNext
    Loop
    rsD.Close
    
    Set RsT = Db.OpenRecordset("Icube_", dbOpenDynaset)
    Do While Not RsT.EOF
        Dim JurisCode As String: JurisCode = Trim(Nz(RsT!施工管轄組織コード, ""))
        If Dict.Exists(JurisCode) Then
            RsT.Edit
            RsT!施工管轄組織名 = Dict(JurisCode)
            RsT.Update
        End If
        RsT.MoveNext
    Loop
    RsT.Close
End Sub

Private Sub Process_Split_ProjectNames()
    Dim rs As DAO.Recordset: Set rs = CurrentDb.OpenRecordset("Icube_", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!一件工事判定 = "小口工事" Then
            Dim ProjName As String: ProjName = Nz(rs!基本工事名称, "")
            rs.Edit
            ' 作業所判定
            If Left(ProjName, 3) = "建築部" Then
                rs!基本工事名_作業所 = "建築部"
            Else
                Dim PosRN As Long: PosRN = InStr(ProjName, "ＲＮ")
                If PosRN >= 3 Then
                    rs!基本工事名_作業所 = Mid(ProjName, PosRN - 2, 2)
                Else
                    rs!基本工事名_作業所 = ""
                End If
            End If
            ' 年度・Q・官民・繰越
            Dim PosYear As Long: PosYear = InStr(ProjName, "年度")
            rs!基本工事名_年度 = IIf(PosYear >= 3, Mid(ProjName, PosYear - 2, 2), "")
            rs!基本工事名_Q = IIf(PosYear > 0 And Len(ProjName) >= PosYear + 2, Mid(ProjName, PosYear + 2, 2), "")
            rs!基本工事名_官民 = IIf(InStr(ProjName, "民間") > 0, "民間", IIf(InStr(ProjName, "官庁") > 0, "官庁", ""))
            rs!基本工事名_繰越 = IIf(InStr(ProjName, "（繰越）") > 0, "（繰越）", "")
            rs.Update
        Else
            rs.Edit
            rs!基本工事名_作業所 = "N/A": rs!基本工事名_年度 = "N/A": rs!基本工事名_Q = "N/A": rs!基本工事名_官民 = "N/A": rs!基本工事名_繰越 = "N/A"
            rs.Update
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub Process_Calculate_PeriodFromName()
    CurrentDb.Execute "UPDATE Icube_ SET 基本工事名_期 = Val(StrConv(Nz(基本工事名_年度,''), 8)) - 12 WHERE 基本工事名_年度 <> '' AND 基本工事名_年度 <> 'N/A'", dbFailOnError
End Sub

Private Sub Process_Transfer_TempProjectCode()
    Call Internal_MapTempProject("仮基本工事コード", "仮基本工事略名", False)
End Sub

Private Sub Process_Map_OrderFieldsToIcube()
    Call Internal_MapTempProject("仮基本工事コード_受注", "仮基本工事略名_受注", True)
End Sub

Private Sub Internal_MapTempProject(ByVal FldC As String, ByVal FldN As String, ByVal IsO As Boolean)
    Dim RsI As DAO.Recordset, RsR As DAO.Recordset: Set RsR = CurrentDb.OpenRecordset("tb_仮基本工事", dbOpenSnapshot)
    Set RsI = CurrentDb.OpenRecordset("SELECT * FROM Icube_ WHERE 一件工事判定 = '小口工事'", dbOpenDynaset)
    Do While Not RsI.EOF
        Dim wS As String: wS = Trim(Nz(RsI!基本工事名_作業所, ""))
        Dim pP As String: pP = Trim(Nz(RsI!基本工事名_官民, ""))
        Dim qV As String: qV = IIf(IsO, ConvertToZenkakuNumber(Trim(Nz(RsI!受注Q, ""))) & "Q", Trim(Nz(RsI!基本工事名_Q, "")))
        Dim cR As String: cR = Trim(Nz(RsI!基本工事名_繰越, ""))
        RsR.MoveFirst
        Do While Not RsR.EOF
            Dim Match As Boolean: Match = (Trim(Nz(RsR!基本工事名_作業所, "")) = wS And Trim(Nz(RsR!基本工事名_官民, "")) = pP And Trim(Nz(RsR!基本工事名_Q, "")) = qV)
            If Not IsO Then
                Match = Match And (Trim(Nz(RsR!基本工事名_繰越, "")) = cR)
            End If
            If Match Then
                RsI.Edit
                RsI.fields(FldC).Value = RsR!仮基本工事コード
                RsI.fields(FldN).Value = RsR!仮基本工事略名
                RsI.Update: Exit Do
            End If
            RsR.MoveNext
        Loop
        RsI.MoveNext
    Loop
    RsI.Close: RsR.Close
End Sub

'---------------------------------------------------------------------------------------------------
' 3. ヘルパー関数 (Private)
'---------------------------------------------------------------------------------------------------

Private Function GetCleanedName_FromMaster(ByVal CCode As String, ByVal Raw As String) As String
    Dim rs As DAO.Recordset: Dim ResText As String: ResText = Raw
    Set rs = CurrentDb.OpenRecordset("tbl_工事名cle", dbOpenSnapshot)
    Do Until rs.EOF
        If Nz(rs!発注者コード, "") = CCode Then
            Dim Tri As String: Tri = Nz(rs!トリガーワード, "")
            Dim Del As String: Del = Nz(rs!del区分ワード, "")
            If Tri = "" Or Left(ResText, Len(Tri)) = Tri Then
                Dim Pos As Long
                If Del = "ブランク" Then
                    Pos = InStr(ResText, " ")
                    If Pos = 0 Then Pos = InStr(ResText, "　")
                    If Pos > 0 Then ResText = Mid(ResText, Pos + 1)
                Else
                    Pos = InStr(ResText, Del)
                    If Pos > 0 Then ResText = Mid(ResText, Pos + Len(Del))
                End If
                ResText = Trim(ResText)
                Exit Do
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close: GetCleanedName_FromMaster = ResText
End Function

Private Function GetFiscalYear(ByVal dt As Date) As Integer
    GetFiscalYear = IIf(VBA.Month(dt) <= 3, VBA.Year(dt) - 1, VBA.Year(dt))
End Function

Private Function GetFiscalQuarter(ByVal dt As Date) As Integer
    Dim m As Integer: m = VBA.Month(dt)
    Select Case m
        Case 4 To 6: GetFiscalQuarter = 1
        Case 7 To 9: GetFiscalQuarter = 2
        Case 10 To 12: GetFiscalQuarter = 3
        Case 1 To 3: GetFiscalQuarter = 4
    End Select
End Function

Private Function ConvertToZenkakuNumber(ByVal s As String) As String
    Dim i As Integer, c As String, r As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c Like "[0-9]" Then
            r = r & Chr(Asc(c) - Asc("0") + &H824F)
        Else
            r = r & c
        End If
    Next
    ConvertToZenkakuNumber = r
End Function
