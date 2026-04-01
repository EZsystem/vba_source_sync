Attribute VB_Name = "acc_mod_Icube_Validator"
'Attribute VB_Name = "acc_mod_Icube_Validator"
Option Compare Database
Option Explicit

'===================================================================================================
' モジュール名 : acc_mod_Icube_Validator
' 概要         : at_Icube テーブル固有のバリデーション・判定・補完ロジック
' 依存関係     : acc_clsDataCleaner, com_clsErrorUtility, acc_mod_MappingTemplate
'===================================================================================================

' 会計年度計算用の定数
Private Const BASE_YEAR As Integer = 2012

'---------------------------------------------------------------------------------------------------
' 1. 公開プロシージャ (司令塔)
'---------------------------------------------------------------------------------------------------

' Phase 1-2
Public Sub Process_BasicValidation_And_Split(ByRef cleaner As acc_clsDataCleaner, ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Call Process_Judge_OneTimeProject
    Call Process_Copy_Empty_ProjectInfo
    Call Process_Merge_BranchCode
    Call Process_DateConversion_Smart(cleaner)
    Call Process_Update_Jurisdiction(ErrorLog)
    Call Process_Split_ProjectNames
    Call Process_Calculate_PeriodFromName
    Call Process_Transfer_TempProjectCode
    Call Process_Map_OrderFieldsToIcube
    Exit Sub
Err_Handler:
    ErrorLog.Notify_Smart_Popup "Validator Phase 1-2 Error: " & Err.Description
End Sub

' Phase 3-4
Public Sub Process_Category_And_Price(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim DbObj As DAO.Database: Set DbObj = CurrentDb
    Dim RsMain As DAO.Recordset, rsMap As DAO.Recordset
    Dim RawText As String, ProjectPrice As Currency

    ' Phase 3: 用途補正 (OpenRecordset に直接渡す場合は括弧なし)
    Set RsMain = DbObj.OpenRecordset(AT_ICUBE, dbOpenDynaset)
    Set rsMap = DbObj.OpenRecordset(AT_BUILDING_USE_MAP, dbOpenSnapshot)
    Do While Not RsMain.EOF
        RawText = Trim(Nz(RsMain!用途大区分, ""))
        rsMap.MoveFirst
        Do While Not rsMap.EOF
            If RawText = Trim(Nz(rsMap!誤_用途大区分, "")) Then
                RsMain.Edit
                RsMain!s用途大区分 = Trim(Nz(rsMap!正_用途大区分, ""))
                RsMain!s用途大区分名 = Trim(Nz(rsMap!正_用途大区分名, ""))
                RsMain.Update
                Exit Do
            End If
            rsMap.MoveNext
        Loop
        RsMain.MoveNext
    Loop
    RsMain.Close

    ' Phase 4: 金額区分
    Set RsMain = DbObj.OpenRecordset(AT_ICUBE, dbOpenDynaset)
    Set rsMap = DbObj.OpenRecordset(AT_PRICE_CATEGORY_MAP, dbOpenSnapshot)
    Do While Not RsMain.EOF
        ProjectPrice = CCur(Nz(RsMain!工事価格, 0))
        rsMap.MoveFirst
        Do While Not rsMap.EOF
            If ProjectPrice >= CCur(Nz(rsMap!最小金額, 0)) And ProjectPrice <= CCur(Nz(rsMap!最大金額, 0)) Then
                RsMain.Edit
                RsMain!工事金額区分コード = rsMap!工事金額区分コード
                RsMain!工事金額区分名 = rsMap!工事金額区分名
                RsMain!工事金額マイナス判定 = rsMap!工事金額マイナス判定
                RsMain.Update
                Exit Do
            End If
            rsMap.MoveNext
        Loop
        RsMain.MoveNext
    Loop
    RsMain.Close
    Exit Sub
Err_Handler:
    ErrorLog.Notify_Smart_Popup "Validator Phase 3-4 Error: " & Err.Description
End Sub

' Phase 5-6
Public Sub Process_Transcribe_ProjectInfo(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim DbObj As DAO.Database: Set DbObj = CurrentDb
    Dim RsT As DAO.Recordset
    Dim PCode As String
    
    ' Execute(SQL) 内では角括弧を維持
    DbObj.Execute "UPDATE [" & AT_ICUBE & "] SET s基本工事コード = 基本工事コード, s基本工事名称 = 基本工事名称 WHERE 基本工事コード IS NOT NULL", dbFailOnError
    
    ' SQL形式でのOpenRecordsetは角括弧が必要
    Set RsT = DbObj.OpenRecordset("SELECT No, s基本工事コード, 工事コード, 工事帳票名, s基本工事名称 FROM [" & AT_ICUBE & "]", dbOpenDynaset)
    Do While Not RsT.EOF
        PCode = Trim(UCase(Nz(RsT!s基本工事コード, "")))
        If Not PCode Like "KT*" Then
            RsT.Edit
            RsT!s基本工事コード = RsT!工事コード
            RsT!s基本工事名称 = RsT!工事帳票名
            RsT.Update
        End If
        RsT.MoveNext
    Loop
    RsT.Close: Exit Sub
Err_Handler:
    ErrorLog.Notify_Smart_Popup "Validator Phase 5-6 Error: " & Err.Description
End Sub

' Phase 7-8
Public Sub Process_Final_Cleansing(ByRef ErrorLog As com_clsErrorUtility)
    On Error GoTo Err_Handler
    Dim RsT As DAO.Recordset, rsS As DAO.Recordset
    Dim Org As String, Cln As String, CCode As String
    
    Set RsT = CurrentDb.OpenRecordset("SELECT No, 発注者コード, 追加工事名称, 追加工事名称_cle FROM [" & AT_ICUBE & "]", dbOpenDynaset)
    Do While Not RsT.EOF
        Org = Nz(RsT!追加工事名称, "")
        Cln = GetCleanedName_FromMaster(Nz(RsT!発注者コード, ""), Org)
        Cln = Replace(Replace(Replace(Cln, " ", ""), "　", ""), vbTab, "")
        
        Do While InStr(Cln, "【") > 0 And InStr(Cln, "】") > InStr(Cln, "【")
            Cln = Left(Cln, InStr(Cln, "【") - 1) & Mid(Cln, InStr(Cln, "】") + 1)
        Loop
        
        RsT.Edit
        RsT!追加工事名称_cle = Replace(StrConv(Cln, vbWide), "??", "(株)")
        RsT.Update
        RsT.MoveNext
    Loop
    RsT.Close

    Set rsS = CurrentDb.OpenRecordset(AT_CLIENT_DATA, dbOpenSnapshot)
    Do While Not rsS.EOF
        CCode = Nz(rsS!顧客コード, "")
        If CCode <> "" Then
            Set RsT = CurrentDb.OpenRecordset("SELECT 発注者名_tbl FROM [" & AT_ICUBE & "] WHERE 発注者コード = '" & CCode & "' AND (発注者名_tbl IS NULL OR 発注者名_tbl = '')", dbOpenDynaset)
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
    rsS.Close: Exit Sub
Err_Handler:
    ErrorLog.Notify_Smart_Popup "Validator Phase 7-8 Error: " & Err.Description
End Sub

'---------------------------------------------------------------------------------------------------
' 2. 内部補助 (Private)
'---------------------------------------------------------------------------------------------------

Private Sub Process_Judge_OneTimeProject()
    Dim rs As DAO.Recordset: Set rs = CurrentDb.OpenRecordset(AT_ICUBE, dbOpenDynaset)
    Dim Conds As Variant: Conds = Array("１２諸工事", "１３諸工事", "１Ｑ", "２Ｑ", "３Ｑ", "４Ｑ")
    Dim i As Integer, IsSmall As Boolean
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
    CurrentDb.Execute "UPDATE [" & AT_ICUBE & "] SET 基本工事コード = 工事コード WHERE 基本工事コード IS NULL OR 基本工事コード = 'N/A'", dbFailOnError
    CurrentDb.Execute "UPDATE [" & AT_ICUBE & "] SET 基本工事名称 = 工事帳票名 WHERE 基本工事名称 IS NULL OR 基本工事名称 = '' OR 基本工事名称 = 'N/A'", dbFailOnError
End Sub

Private Sub Process_Merge_BranchCode()
    CurrentDb.Execute "UPDATE [" & AT_ICUBE & "] SET 枝番工事コード = Nz(工事コード,'') & '-' & Nz(工事枝番,'')", dbFailOnError
End Sub

Private Sub Process_DateConversion_Smart(ByRef cleaner As acc_clsDataCleaner)
    Dim rs As DAO.Recordset
    Dim Flds As Variant: Flds = Array("[データ年月（受注計上年月）]", "[完成年月日（枝番単位）]")
    Dim Prfx As Variant: Prfx = Array("受注", "完工")
    Dim i As Integer, TargetFld As String, TDate As Date
    
    For i = 0 To UBound(Flds)
        TargetFld = IIf(Prfx(i) = "受注", "受注計上日_日付型", Prfx(i) & "日_日付型")
        Set rs = CurrentDb.OpenRecordset("SELECT No, " & Flds(i) & ", " & Prfx(i) & "年度, " & Prfx(i) & "期, " & Prfx(i) & "Q, " & Prfx(i) & "月, " & TargetFld & ", 一件工事判定, 基本工事名称 FROM [" & AT_ICUBE & "]", dbOpenDynaset)
        Do While Not rs.EOF
            TDate = cleaner.CleanDate(rs.Fields(1).Value)
            If Year(TDate) > 1900 Then
                rs.Edit
                rs.Fields(Prfx(i) & "年度").Value = GetFiscalYear(TDate)
                
                ' --- 完工期かつ小口工事の場合のみ名称から算出 ---
                Dim isSmallHandled As Boolean: isSmallHandled = False
                If Prfx(i) = "完工" And rs!一件工事判定 = "小口工事" Then
                    Dim projName As String: projName = Nz(rs!基本工事名称, "")
                    Dim posYear As Long: posYear = InStr(projName, "年度")
                    If posYear >= 3 Then
                        Dim yearVal As Integer
                        yearVal = val(StrConv(Mid(projName, posYear - 2, 2), vbNarrow))
                        rs.Fields(Prfx(i) & "期").Value = yearVal - 12
                        isSmallHandled = True
                    End If
                End If
                
                ' 通常ロジック (一件工事、または受注期の場合)
                If Not isSmallHandled Then
                    rs.Fields(Prfx(i) & "期").Value = GetFiscalYear(TDate) - BASE_YEAR + 1
                End If
                
                rs.Fields(Prfx(i) & "Q").Value = GetFiscalQuarter(TDate)
                rs.Fields(Prfx(i) & "月").Value = Month(TDate)
                rs.Fields(TargetFld).Value = TDate
                rs.Update
            End If
            rs.MoveNext
        Loop
        rs.Close
    Next i
End Sub

Private Sub Process_Update_Jurisdiction(ByRef ErrorLog As com_clsErrorUtility)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsD As DAO.Recordset, RsT As DAO.Recordset, RsE As DAO.Recordset, dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Set rsD = db.OpenRecordset("SELECT 組織コード, 施工管轄組織コード FROM [" & AT_BRANCH_WORK_HISTORY & "]", dbOpenSnapshot)
    Do While Not rsD.EOF
        If Not IsNull(rsD!組織コード) Then dict(Trim(CStr(rsD!組織コード))) = rsD!施工管轄組織コード
        rsD.MoveNext
    Loop
    rsD.Close
    
    Set RsT = db.OpenRecordset(AT_ICUBE, dbOpenDynaset)
    Set RsE = db.OpenRecordset(AT_ERR_SAGYOSHO, dbOpenDynaset)
    Do While Not RsT.EOF
        Dim oC As String: oC = Trim(Nz(RsT!施工担当組織コード, ""))
        If dict.Exists(oC) Then
            RsT.Edit
            RsT!施工管轄組織コード = dict(oC)
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
    
    Set rsD = db.OpenRecordset("SELECT 施工管轄組織コード, 施工管轄組織名 FROM [" & AT_JURISDICTION_MAP & "]", dbOpenSnapshot)
    dict.RemoveAll
    Do While Not rsD.EOF
        If Not IsNull(rsD!施工管轄組織コード) Then dict(Trim(CStr(rsD!施工管轄組織コード))) = rsD!施工管轄組織名
        rsD.MoveNext
    Loop
    rsD.Close
    
    Set RsT = db.OpenRecordset(AT_ICUBE, dbOpenDynaset)
    Do While Not RsT.EOF
        Dim JC As String: JC = Trim(Nz(RsT!施工管轄組織コード, ""))
        If dict.Exists(JC) Then
            RsT.Edit
            RsT!施工管轄組織名 = dict(JC)
            RsT.Update
        End If
        RsT.MoveNext
    Loop
    RsT.Close
End Sub

Private Sub Process_Split_ProjectNames()
    Dim rs As DAO.Recordset: Set rs = CurrentDb.OpenRecordset(AT_ICUBE, dbOpenDynaset)
    Do While Not rs.EOF
        If rs!一件工事判定 = "小口工事" Then
            Dim Proj As String: Proj = Nz(rs!基本工事名称, "")
            rs.Edit
            If Left(Proj, 3) = "建築部" Then
                rs!基本工事名_作業所 = "建築部"
            Else
                Dim PosRN As Long: PosRN = InStr(Proj, "ＲＮ")
                rs!基本工事名_作業所 = IIf(PosRN >= 3, Mid(Proj, PosRN - 2, 2), "")
            End If
            
            Dim PY As Long: PY = InStr(Proj, "年度")
            rs!基本工事名_年度 = IIf(PY >= 3, Mid(Proj, PY - 2, 2), "")
            rs!基本工事名_Q = IIf(PY > 0 And Len(Proj) >= PY + 2, Mid(Proj, PY + 2, 2), "")
            rs!基本工事名_官民 = IIf(InStr(Proj, "民間") > 0, "民間", IIf(InStr(Proj, "官庁") > 0, "官庁", ""))
            rs!基本工事名_繰越 = IIf(InStr(Proj, "（繰越）") > 0, "（繰越）", "")
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

Public Sub Process_Calculate_PeriodFromName()
    CurrentDb.Execute "UPDATE [" & AT_ICUBE & "] SET 基本工事名_期 = Val(StrConv(Nz(基本工事名_年度,''), 8)) - 12 WHERE 基本工事名_年度 <> '' AND 基本工事名_年度 <> 'N/A'", dbFailOnError
End Sub

Public Sub Process_Transfer_TempProjectCode()
    Call Internal_MapTempProject("仮基本工事コード", "仮基本工事略名", False)
End Sub

Public Sub Process_Map_OrderFieldsToIcube()
    Call Internal_MapTempProject("仮基本工事コード_受注", "仮基本工事略名_受注", True)
End Sub

Private Sub Internal_MapTempProject(ByVal FldC As String, ByVal FldN As String, ByVal IsO As Boolean)
    Dim RsI As DAO.Recordset, RsR As DAO.Recordset
    Set RsR = CurrentDb.OpenRecordset(AT_TEMP_PROJECT_MAP, dbOpenSnapshot)
    Set RsI = CurrentDb.OpenRecordset("SELECT * FROM [" & AT_ICUBE & "] WHERE 一件工事判定 = '小口工事'", dbOpenDynaset)
    
    Do While Not RsI.EOF
        Dim ws As String: ws = Trim(Nz(RsI!基本工事名_作業所, ""))
        Dim pP As String: pP = Trim(Nz(RsI!基本工事名_官民, ""))
        Dim qV As String: qV = IIf(IsO, ConvertToZenkakuNumber(Trim(Nz(RsI!受注Q, ""))) & "Q", Trim(Nz(RsI!基本工事名_Q, "")))
        Dim cR As String: cR = Trim(Nz(RsI!基本工事名_繰越, ""))
        
        RsR.MoveFirst
        Do While Not RsR.EOF
            Dim Match As Boolean: Match = (Trim(Nz(RsR!基本工事名_作業所, "")) = ws And Trim(Nz(RsR!基本工事名_官民, "")) = pP And Trim(Nz(RsR!基本工事名_Q, "")) = qV)
            If Not IsO Then Match = Match And (Trim(Nz(RsR!基本工事名_繰越, "")) = cR)
            
            If Match Then
                RsI.Edit
                RsI.Fields(FldC).Value = RsR!仮基本工事コード
                RsI.Fields(FldN).Value = RsR!仮基本工事略名
                RsI.Update
                Exit Do
            End If
            RsR.MoveNext
        Loop
        RsI.MoveNext
    Loop
    RsI.Close: RsR.Close
End Sub

Private Function GetCleanedName_FromMaster(ByVal CCode As String, ByVal Raw As String) As String
    Dim rs As DAO.Recordset: Dim ResText As String: ResText = Raw
    ' 修正: 直接テーブル名として開くため括弧を除去
    Set rs = CurrentDb.OpenRecordset(AT_PROJECT_NAME_CLEAN, dbOpenSnapshot)
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
    GetFiscalYear = IIf(Month(dt) <= 3, Year(dt) - 1, Year(dt))
End Function

Private Function GetFiscalQuarter(ByVal dt As Date) As Integer
    Select Case Month(dt)
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

