Attribute VB_Name = "mod_tabCopy2"
Option Compare Database
'Icube_から関係テーブルへの転写
Public Sub mod_tabCopy2_Aall()
    ' A-1: 基本工事_完工
    Call mod_tabCopy2_A_kihonKanko
    ' A-2: 基本工事_作業所
    Call mod_tabCopy2_A_kihonSagyo
    ' A-3: 基本工事_受注
    Call mod_tabCopy2_A_kihonJyuchu
    ' A-4: 工事コード情報
    Call mod_tabCopy2_A_kojicode
    ' A-5: 枝番工事
    Call mod_tabCopy2_A_edaban
End Sub


' A-1: 基本工事_完工
Public Sub mod_tabCopy2_A_kihonKanko()
    Call TransferTable("Icube_", "kt_基本工事_完工", "基本工事コード")
End Sub

' A-2: 基本工事_作業所
Public Sub mod_tabCopy2_A_kihonSagyo()
    Call TransferTable("Icube_", "kt_基本工事_作業所", "基本工事コード")
End Sub

' A-3: 基本工事_受注
Public Sub mod_tabCopy2_A_kihonJyuchu()
    Call TransferTable("Icube_", "kt_基本工事_受注", "基本工事コード")
End Sub

' A-4: 工事コード情報
Public Sub mod_tabCopy2_A_kojicode()
    Call TransferTable("Icube_", "kt_工事コード情報", "工事コード")
End Sub

' A-5: 枝番工事
Public Sub mod_tabCopy2_A_edaban()
    Call TransferTable("Icube_", "kt_枝番工事", "枝番工事コード")
End Sub

'Icube_累計から関係テーブルへの転写
Public Sub mod_tabCopy2Ball()
    ' B-1: 基本工事_完工
    Call mod_tabCopy2_B_kihonKanko
    ' B-2: 基本工事_作業所
    Call mod_tabCopy2_B_kihonSagyo
    ' B-3: 基本工事_受注
    Call mod_tabCopy2_B_kihonJyuchu
    ' B-4: 工事コード情報
    Call mod_tabCopy2_B_kojicode
    ' B-5: 枝番工事
    Call mod_tabCopy2_B_edaban
End Sub




' B-1: 基本工事_完工（累計）
Public Sub mod_tabCopy2_B_kihonKanko()
    Call TransferTable("Icube_累計", "kt_基本工事_完工", "基本工事コード")
End Sub

' B-2: 基本工事_作業所（累計）
Public Sub mod_tabCopy2_B_kihonSagyo()
    Call TransferTable("Icube_累計", "kt_基本工事_作業所", "基本工事コード")
End Sub

' B-3: 基本工事_受注（累計）
Public Sub mod_tabCopy2_B_kihonJyuchu()
    Call TransferTable("Icube_累計", "kt_基本工事_受注", "基本工事コード")
End Sub

' B-4: 工事コード情報（累計）
Public Sub mod_tabCopy2_B_kojicode()
    Call TransferTable("Icube_累計", "kt_工事コード情報", "工事コード")
End Sub

' B-5: 枝番工事（累計）
Public Sub mod_tabCopy2_B_edaban()
    Call TransferTable("Icube_累計", "kt_枝番工事", "枝番工事コード")
End Sub



Public Sub TransferTable( _
    SourceTable As String, _
    targetTable As String, _
    conditionField As String _
)
    On Error GoTo ErrHandle

    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim condValue As Variant
    Dim fld As DAO.Field

    Set db = CurrentDb()
    Set rsSource = db.OpenRecordset("SELECT * FROM [" & SourceTable & "]", dbOpenSnapshot)

Set rsTarget = db.OpenRecordset(targetTable, dbOpenDynaset)


    Do While Not rsSource.EOF
        condValue = rsSource(conditionField)

        If DCount("*", "[" & targetTable & "]", "[" & conditionField & "] = '" & replace(condValue, "'", "''") & "'") = 0 Then
            rsTarget.AddNew

            For Each fld In rsSource.Fields
                If FieldExists(rsTarget, fld.Name) Then
                    On Error Resume Next
                    rsTarget(fld.Name).Value = fld.Value
                    If Err.Number <> 0 Then
                        Debug.Print "【型エラー】フィールド名: " & fld.Name
                        Debug.Print "→ 値: " & fld.Value & " ／ 型: " & typeName(fld.Value)
                        Debug.Print "→ エラー内容: " & Err.description
                        Err.Clear
                    End If
                    On Error GoTo ErrHandle
                End If
            Next fld

            rsTarget.Update
        End If

        rsSource.MoveNext
    Loop

    'MsgBox "転写完了にゃ！ [" & sourceTable & "] → [" & targetTable & "]", vbInformation
    GoTo Finalize

Finalize:
    On Error Resume Next
    rsSource.Close: Set rsSource = Nothing
    rsTarget.Close: Set rsTarget = Nothing
    Set db = Nothing
    Exit Sub

ErrHandle:
    MsgBox "エラーが発生したニャ：" & Err.description, vbCritical
    Resume Finalize
End Sub


