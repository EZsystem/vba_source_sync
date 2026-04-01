Attribute VB_Name = "acc_mod_VbaSync"
'Attribute VB_Name = "acc_mod_VbaSync"
Option Compare Database
Option Explicit

'================================================================
' Module: acc_mod_VbaSync
' 説明   : 外部フォルダから .bas / .cls ファイルを一括インポートする同期ツール
' 更新日 : 2026/04/01
'================================================================

'----------------------------------------------------------------
' プロシージャ名 : Sync_Vba_Project
' 概要           : ワークスペースのソースファイルをAccessプロジェクトに反映します。
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
    ' ※ 処理名称 "VBAソースコード同期" のレコードを参照します
    ' SQL内の名称は、テーブルに登録する名前と一致させてください
    Set rs = db.OpenRecordset("SELECT [既定パス] FROM [_at_SystemRegistry] WHERE [処理名称] = 'VBAソースコード同期'", dbOpenSnapshot)
    
    If rs.EOF Then
        MsgBox "システムレジストリに 'VBAソースコード同期' の設定が見つかりません。" & vbCrLf & _
               "[_at_SystemRegistry] テーブルにレコードを追加してください。", vbCritical
        Exit Sub
    End If
    
    folderPath = Nz(rs![既定パス], "")
    rs.Close
    
    ' 2. パスの妥当性チェック
    If folderPath = "" Then
        MsgBox "VBA同期用のパスが空欄です。[_at_SystemRegistry] の [既定パス] を設定してください。", vbExclamation
        Exit Sub
    End If
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "指定されたフォルダが見つかりません：" & vbCrLf & folderPath, vbCritical
        Exit Sub
    End If
    
    ' 3. 更新対象ファイルのリストアップ
    fileList = ""
    For Each fileObj In fso.GetFolder(folderPath).Files
        Select Case LCase(fso.GetExtensionName(fileObj.Name))
            Case "bas", "cls"
                ' 自分自身は除外（自己破壊防止）
                compName = fso.GetBaseName(fileObj.Name)
                If compName <> "acc_mod_VbaSync" Then
                    fileList = fileList & " - " & fileObj.Name & vbCrLf
                End If
        End Select
    Next
    
    If fileList = "" Then
        MsgBox "対象フォルダに .bas または .cls ファイルが見つかりません。", vbInformation
        Exit Sub
    End If
    
    ' 4. 実行前確認（Ezさんのご要望：確認メッセージを表示）
    If MsgBox("以下のモジュールをワークスペースからインポート（上書き更新）しますか？" & vbCrLf & _
              "※既存の同名モジュールは削除されます。" & vbCrLf & vbCrLf & _
              fileList, vbQuestion + vbOKCancel, "VBA同期の実行確認") <> vbOK Then
        Exit Sub
    End If
    
    ' 5. インポートの実行
    ' Application.VBE オブジェクトを使用してプロジェクトを操作します
    Dim vbeProj As Object
    On Error Resume Next
    Set vbeProj = Application.VBE.ActiveVBProject
    If Err.Number <> 0 Then
        MsgBox "VBAプロジェクトへのアクセスに失敗しました。" & vbCrLf & _
               "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」設定を確認してください。", vbCritical
        Exit Sub
    End If
    On Error GoTo Err_Handler
    
    For Each fileObj In fso.GetFolder(folderPath).Files
        Select Case LCase(fso.GetExtensionName(fileObj.Name))
            Case "bas", "cls"
                compName = fso.GetBaseName(fileObj.Name)
                
                ' 自己保護（自身は更新しない）
                If compName <> "acc_mod_VbaSync" Then
                    ' 既存コンポーネントの削除
                    On Error Resume Next
                    vbeProj.VBComponents.Remove vbeProj.VBComponents(compName)
                    On Error GoTo Err_Handler
                    
                    ' ファイルのインポート
                    vbeProj.VBComponents.Import fileObj.Path
                    updateCount = updateCount + 1
                End If
        End Select
    Next
    
    MsgBox updateCount & " 件のモジュールを最新状態に同期しました。", vbInformation
    Exit Sub

Err_Handler:
    MsgBox "同期処理中にエラーが発生しました：" & vbCrLf & Err.Description, vbCritical
End Sub


