'Attribute VB_Name = "acc_mod_VbaSync"
Option Compare Database
Option Explicit

'================================================================
' Module: acc_mod_VbaSync
' 説明   : 外部フォルダから .bas / .cls ファイルを一括インポートする同期ツール
' 更新日 : 2026/04/01 (快適版：削除時・保存時の確認を極力自動化)
'================================================================

Public Sub Sync_Vba_Project()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim fso As Object
    Dim folderPath As String
    Dim fileObj As Object
    Dim fileList As String
    Dim updateCount As Long
    
    On Error GoTo Err_Handler
    
    ' 1. パス取得
    Set rs = db.OpenRecordset("SELECT [既定パス] FROM [_at_SystemRegistry] WHERE [処理名称] = 'VBAソースコード同期'", dbOpenSnapshot)
    If rs.EOF Then MsgBox "システムレジストリの設定なし", vbCritical: Exit Sub
    folderPath = Nz(rs![既定パス], ""): rs.Close
    If folderPath = "" Then Exit Sub
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 2. インポート対象の確認
    fileList = ""
    For Each fileObj In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(fileObj.Name)) Like "[bc][al][ss]" Then
            If fso.GetBaseName(fileObj.Name) <> "acc_mod_VbaSync" Then
                fileList = fileList & " - " & fileObj.Name & vbCrLf
            End If
        End If
    Next
    
    If MsgBox("【高速同期】モジュールを全入替し、最後に一括保存・コンパイルを行います。" & vbCrLf & _
              "※途中の保存確認ダイアログは自動的にスキップされます。", vbInformation + vbOKCancel, "VBA同期の実行（快適版）") <> vbOK Then
        Exit Sub
    End If
    
    ' 3. 強力なクリーンアップ（退避リネーム削除）
    Dim vbeProj As Object: Set vbeProj = Application.VBE.ActiveVBProject
    Dim i As Long
    
    ' リセット実行（実行ロックの解除）
    On Error Resume Next
    Application.VBE.CommandBars.FindControl(ID:=228).Execute
    On Error GoTo Err_Handler
    
    For i = vbeProj.VBComponents.count To 1 Step -1
        Dim comp As Object: Set comp = vbeProj.VBComponents(i)
        If comp.Name <> "acc_mod_VbaSync" And (comp.Type = 1 Or comp.Type = 2) Then
            ' 【ポイント】削除時の確認ダイアログを出さないための処理
            Call AtomicRemove_Silent(comp)
        End If
    Next i
    
    ' 4. インポートの実行
    For Each fileObj In fso.GetFolder(folderPath).Files
        Dim ext As String: ext = LCase(fso.GetExtensionName(fileObj.Name))
        If (ext = "bas" Or ext = "cls") And fso.GetBaseName(fileObj.Name) <> "acc_mod_VbaSync" Then
            Dim targetName As String: targetName = fso.GetBaseName(fileObj.Name)
            ' 重複の残骸があれば執拗に消去
            On Error Resume Next
            Call AtomicRemove_Silent(vbeProj.VBComponents(targetName))
            Call AtomicRemove_Silent(vbeProj.VBComponents(targetName & "1"))
            On Error GoTo Err_Handler
            
            Call ImportFromUtf8(vbeProj, fileObj.Path)
            updateCount = updateCount + 1
        End If
    Next
    
    ' 5. 【仕上げ】全モジュールの一括保存とコンパイル（ダイアログ防止の決定打）
    On Error Resume Next
    DoCmd.RunCommand acCmdCompileAndSaveAllModules
    On Error GoTo Err_Handler
    
    MsgBox updateCount & " 件の同期、および一括保存・コンパイルが完了しました。", vbInformation
    Exit Sub

Err_Handler:
    MsgBox "同期エラー：" & Err.Description, vbCritical
End Sub

' ダイアログを出さずに削除する
Private Sub AtomicRemove_Silent(ByRef comp As Object)
    If comp Is Nothing Then Exit Sub
    On Error Resume Next
    ' 名前を変えて衝突を回避
    comp.Name = "tmp_" & Format(Now, "hhnnss") & "_" & comp.Name
    ' 削除（確認を抑制）
    Application.VBE.ActiveVBProject.VBComponents.Remove comp
    On Error GoTo 0
End Sub

' UTF-8変換インポート（一時ファイルの拡張子を厳守）
Private Sub ImportFromUtf8(ByRef vbeProj As Object, ByVal utf8Path As String)
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    Dim sjisStream As Object: Set sjisStream = CreateObject("ADODB.Stream")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ext As String: ext = fso.GetExtensionName(utf8Path)
    Dim tempPath As String: tempPath = fso.GetParentFolderName(utf8Path) & "\_sync_tmp_" & fso.GetBaseName(utf8Path) & "." & ext
    
    stream.Type = 2: stream.Charset = "UTF-8": stream.Open: stream.LoadFromFile utf8Path
    sjisStream.Type = 2: sjisStream.Charset = "shift-jis": sjisStream.Open: stream.CopyTo sjisStream
    
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
    sjisStream.SaveToFile tempPath, 2
    stream.Close: sjisStream.Close
    
    vbeProj.VBComponents.Import tempPath
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
End Sub
