Attribute VB_Name = "mod_com"
Option Compare Database
Option Explicit


'参照ライブラリーをデバッグプリントに出力する
Public Sub mod_com1()
    Dim ref As Reference
    
    For Each ref In Application.References
        Debug.Print ref.Name
    Next ref
End Sub


'-------------------------------------------------
' ■ フィールドが転写先に存在するかチェック
'-------------------------------------------------

Public Function FieldExists(rs As DAO.Recordset, FieldName As String) As Boolean
    On Error Resume Next
    FieldExists = Not IsNull(rs.fields(FieldName).Name)
    On Error GoTo 0
End Function
