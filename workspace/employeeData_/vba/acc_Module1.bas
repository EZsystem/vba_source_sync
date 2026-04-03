Attribute VB_Name = "Module1"
'--------------------------------------------
' プロシージャ名： Generate_Kehi_History_Data_V3
' 概要： ID(手動数値型)を自動採番しながら過去データを生成する
'--------------------------------------------
Public Sub Generate_Kehi_History_Data_V3()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim i As Long, m As Long, nextID As Long
    Dim targetDate As Date
    Dim qNum As Integer, qFull As String, qHalf As String
    Dim finalName As String, finalCode As String
    
    ' 現在の最大IDを取得して、次の開始IDを決める
    nextID = Nz(DMax("ID", "at_工事経費"), 0) + 1
    
    ' 基準データ（26年度 1Q の基本形）
    Dim baseData As Variant
    baseData = Array( _
        Array("D1M", "山形ＲＮ２６年度１Ｑ民間", "山形ＲＮ（作）"), _
        Array("B1M", "秋田ＲＮ２６年度１Ｑ民間", "秋田ＲＮ（作）"), _
        Array("C1M", "盛岡ＲＮ２６年度１Ｑ民間", "盛岡ＲＮ（作）"), _
        Array("A1M", "青森ＲＮ２６年度１Ｑ民間", "青森ＲＮ（作）"), _
        Array("F1M", "仙台ＲＮ２６年度１Ｑ民間", "仙台ＲＮ（作）"), _
        Array("E1M", "福島ＲＮ２６年度１Ｑ民間", "福島ＲＮ（作）") _
    )
    
    Dim expenseItems As Variant
    expenseItems = Array( _
        Array("車両", 48000), _
        Array("事務所宿舎", 229000), _
        Array("工事費", 159000), _
        Array("雑費", 280000) _
    )
    
    On Error GoTo ErrLine
    DBEngine.BeginTrans
    
    Set rs = db.OpenRecordset("at_工事経費", dbOpenDynaset)
    
    For m = 0 To 11
        targetDate = DateSerial(2025, 4 + m, 1)
        
        ' 四半期判定
        If Month(targetDate) <= 3 Then qNum = 4 Else qNum = (Month(targetDate) - 4) \ 3 + 1
        qFull = StrConv(CStr(qNum) & "Ｑ", vbWide)
        qHalf = CStr(qNum) & "Q"
        
        Dim bIdx As Integer, eIdx As Integer
        For bIdx = 0 To 5
            finalName = baseData(bIdx)(1)
            finalName = Replace(finalName, "２６年度", "２５年度")
            finalName = Replace(finalName, "１Ｑ", qFull)
            
            finalCode = baseData(bIdx)(0)
            Mid(finalCode, 2, 1) = CStr(qNum)
            
            For eIdx = 0 To 3
                rs.AddNew
                rs!ID = nextID  ' ★ここで手動でIDをセット
                rs!年月 = targetDate
                rs!仮基本工事コード = finalCode
                rs!工事名 = finalName
                rs!経費名 = expenseItems(eIdx)(0)
                rs!経費額 = expenseItems(eIdx)(1)
                rs!作業所名 = baseData(bIdx)(2)
                rs!期 = Get_FiscalTerm(targetDate)
                rs!Q = qHalf
                
                rs.Update
                nextID = nextID + 1 ' 次のレコード用に増やす
            Next eIdx
        Next bIdx
    Next m
    
    DBEngine.CommitTrans
    MsgBox "288件のデータを生成しました（開始ID: " & (nextID - 288) & "）", vbInformation
    Exit Sub

ErrLine:
    DBEngine.Rollback
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub


