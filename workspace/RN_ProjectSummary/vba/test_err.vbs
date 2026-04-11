Dim cn, rs, sql
Set cn = CreateObject("ADODB.Connection")
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\My_code\11_workspaces\VBA_manager\vba_source_sync\workspace\RN_ProjectSummary\RN_ProjectSummary.accdb;"
sql = "SELECT [仮基本工事コード] FROM [at_仮基本工事] WHERE [施工管轄組織名] = '福島ＲＮ（作）' AND ([基本工事名_Q] = '4Q' OR [基本工事名_Q] = '4' OR [基本工事名_Q] LIKE '*4*') AND (Nz([基本工事名_官民], '') <> '官庁') AND (Nz([基本工事名_名称分類], '') NOT LIKE '*繰越*') ORDER BY IIf([基本工事名_官民]='民間',0,1), Len([仮基本工事コード]) ASC"
On Error Resume Next
Set rs = cn.Execute(sql)
If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
Else
    if Not rs.EOF Then
        WScript.Echo "Success! First code: " & rs.Fields("仮基本工事コード").Value
    end if
End If
cn.Close
