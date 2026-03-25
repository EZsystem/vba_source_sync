Attribute VB_Name = "mod_Icube_ruikei_copy"
'=============================
' Module: mod_Icube_ruikei_copy
' 説明 : Icube_ に存在する「枝番工事コード」と一致する
'        Icube_累計 のレコードを削除し、再コピーするにゃ
'=============================
Option Compare Database
Option Explicit

Public Sub Refresh_Icube_ruikei_From_Icube_()
    Const TABLE_SRC As String = "Icube_"
    Const TABLE_DEST As String = "Icube_累計"
    Const KEY_FIELD As String = "枝番工事コード"

    Dim srcFetcher As New com_clsAccessFetcher
    Dim destFetcher As New com_clsAccessFetcher
    Dim connector As New acc_clsDbConnector
    Dim srcCodes As Variant, destCodes As Variant
    Dim i As Long, j As Long
    Dim refreshList As Object: Set refreshList = CreateObject("Scripting.Dictionary")

    ' === 転写元コード取得 ===
    srcFetcher.filePath = CurrentProject.FullName
    srcCodes = srcFetcher.FetchArray("SELECT [" & KEY_FIELD & "] FROM [" & TABLE_SRC & "] WHERE [" & KEY_FIELD & "] IS NOT NULL")

    If IsEmpty(srcCodes) Then
        MsgBox "[" & TABLE_SRC & "] に対象レコードがなかったにゃ（処理中止）", vbExclamation
        Exit Sub
    End If

    ' === 転写先コード取得 ===
    destFetcher.filePath = CurrentProject.FullName
    destCodes = destFetcher.FetchArray("SELECT [" & KEY_FIELD & "] FROM [" & TABLE_DEST & "] WHERE [" & KEY_FIELD & "] IS NOT NULL")

    ' === 転写元コードを辞書に保持 ===
    Dim codeMap As Object: Set codeMap = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(srcCodes, 1)
        codeMap(CStr(srcCodes(i, 1))) = True
    Next i

    ' === 重複コードを抽出 ===
    If Not IsEmpty(destCodes) Then
        For j = 1 To UBound(destCodes, 1)
            Dim code As String
            code = CStr(destCodes(j, 1))
            If codeMap.Exists(code) Then
                refreshList(code) = True
            End If
        Next j
    End If

    ' === 削除対象があるなら削除 ===
    connector.Init
    Dim k As Variant
    For Each k In refreshList.Keys
        connector.ExecuteSQL "DELETE FROM [" & TABLE_DEST & "] WHERE [" & KEY_FIELD & "] = '" & replace(k, "'", "''") & "'"
    Next k

    ' === 転写（再INSERT） ===
    For Each k In codeMap.Keys
        connector.ExecuteSQL _
            "INSERT INTO [" & TABLE_DEST & "] SELECT * FROM [" & TABLE_SRC & "] WHERE [" & KEY_FIELD & "] = '" & replace(k, "'", "''") & "'"
    Next k

    MsgBox codeMap.count & " 件のレコードを [" & TABLE_DEST & "] に転写（必要に応じて上書き）したにゃ", vbInformation
End Sub


