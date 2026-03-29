Attribute VB_Name = "acc_mod_LogHelper"
'Attribute VB_Name = "acc_mod_LogHelper"
Option Compare Database
Option Explicit

Public Sub AppendLog(ByVal logMsg As String)
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms("frm_LogConsole")
    If frm Is Nothing Then
        DoCmd.OpenForm "frm_LogConsole", acNormal
        Set frm = Forms("frm_LogConsole")
    End If
    frm!txtLogOutput.Value = Nz(frm!txtLogOutput.Value, "") & Format(Now, "hh:nn:ss") & " > " & logMsg & vbCrLf
    frm!txtLogOutput.SelStart = Len(frm!txtLogOutput.Text)
End Sub

Public Sub ClearLog()
    On Error Resume Next
    Forms("frm_LogConsole")!txtLogOutput.Value = ""
End Sub
