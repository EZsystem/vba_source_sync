Attribute VB_Name = "acc_mod_VbaSync"
'Attribute VB_Name = "acc_mod_VbaSync"
Option Compare Database
Option Explicit

'================================================================
' Module: acc_mod_VbaSync
' 説明   : 外部フォルダから .bas / .cls ファイルを一括インポートする同期ツール
' 更新日 : 2026/04/01 (UTF-8ハイブリッド同期版)
'================================================================

'----------------------------------------------------------------
' プロシージャ名 : Sync_Vba_Project
' 概要           : ワークスペースのUTF-8ソースをShift-JISに自動変換してAccessに同期します。
' 実行条件       : 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」がONであること
'----------------------------------------------------------------
Public Sub Sync_Vba_Project()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim fso As Object
    Dim folderPath As String
    Dim fileObj As Object
    Dim fileList As String
    Dim compName As String
    Dim updateCount As Long
    
    On Error GoTo Err_Handler
    
    ' 1. システムレジストリからパスを取得
    Set rs = db.OpenRecordset("SELECT [既定パス] FROM [_at_SystemRegistry] WHERE [処理名称] = 'VBAソースコード同期'", dbOpenSnapshot)
    
    If rs.EOF Then
        MsgBox "システムレジストリに 'VBAソースコード同期' の設定が見つかりません。", vbCritical
        Exit Sub
    End If
    
    folderPath = Nz(rs![既定パス], "")
    rs.Close
    
    If folderPath = "" Then
        MsgBox "VBA同期用のパスが空欄です。", vbExclamation
        Exit Sub
    End If
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "フォルダが見つかりません：" & vbCrLf & folderPath, vbCritical
        Exit Sub
    End If
    
    ' 2. 更新対象ファイルのリストアップ
    fileList = ""
    For Each fileObj In fso.GetFolder(folderPath).Files
        Select Case LCase(fso.GetExtensionName(fileObj.Name))
            Case "bas", "cls"
                compName = fso.GetBaseName(fileObj.Name)
                ' 自分自身は除外（自己破壊防止）
                If compName <> "acc_mod_VbaSync" Then
                    fileList = fileList & " - " & fileObj.Name & vbCrLf
                End If
        End Select
    Next
    
    If fileList = "" Then
        MsgBox "対象フォルダにソースファイルが見つかりません。", vbInformation
        Exit Sub
    End If
    
    ' 3. 実行前確認
    If MsgBox("以下のモジュールをワークスペースから最新同期（UTF-8変換込）しますか？" & vbCrLf & _
              fileList, vbQuestion + vbOKCancel, "ハイブリッド同期の実行確認") <> vbOK Then
        Exit Sub
    End If
    
    ' 4. インポートの実行
    Dim vbeProj As Object
    On Error Resume Next
    Set vbeProj = Application.VBE.ActiveVBProject
    If Err.Number <> 0 Then
        MsgBox "VBAプロジェクトへのアクセスに失敗しました。「信頼設定」を確認してください。", vbCritical
        Exit Sub
    End If
    On Error GoTo Err_Handler
    
    For Each fileObj In fso.GetFolder(folderPath).Files
        Select Case LCase(fso.GetExtensionName(fileObj.Name))
            Case "bas", "cls"
                compName = fso.GetBaseName(fileObj.Name)
                
                If compName <> "acc_mod_VbaSync" Then
                    ' 既存コンポーネントを削除
                    On Error Resume Next
                    vbeProj.VBComponents.Remove vbeProj.VBComponents(compName)
                    On Error GoTo Err_Handler
                    
                    ' 【ポイント】UTF-8からShift-JISへの一時変換インポートを実行
                    Call ImportFromUtf8(vbeProj, fileObj.Path)
                    updateCount = updateCount + 1
                End If
        End Select
    Next
    
    MsgBox updateCount & " 件のモジュールを、UTF-8から高度同期しました。", vbInformation
    Exit Sub

Err_Handler:
    MsgBox "同期処理中にエラーが発生しました：" & vbCrLf & Err.Description, vbCritical
End Sub

'----------------------------------------------------------------
' 内部補助関数 : ImportFromUtf8
' 概要           : UTF-8ファイルを読み込み、Shift-JISの一時ファイルとしてImportします。
'----------------------------------------------------------------
Private Sub ImportFromUtf8(ByRef vbeProj As Object, ByVal utf8Path As String)
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    Dim sjisStream As Object: Set sjisStream = CreateObject("ADODB.Stream")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tempPath As String: tempPath = utf8Path & ".sync_temp"
    
    ' 1. UTF-8 (BOMなし/あり両対応) で読み込み
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile utf8Path
    
    ' 2. Shift-JIS で別ストリームに書き出し
    sjisStream.Type = 2 ' adTypeText
    sjisStream.Charset = "shift-jis"
    sjisStream.Open
    
    ' 内容をコピー (内部でエンコード変換が行われます)
    stream.CopyTo sjisStream
    
    ' 3. 一時ファイルとして保存
    On Error Resume Next
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
    On Error GoTo 0
    sjisStream.SaveToFile tempPath, 2 ' adSaveCreateOverWrite
    
    stream.Close
    sjisStream.Close
    
    ' 4. インポート実行 (Shift-JISのファイルを読み込ませる)
    vbeProj.VBComponents.Import tempPath
    
    ' 5. 一時ファイルの削除
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
End Sub


