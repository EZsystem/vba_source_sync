Attribute VB_Name = "acc_mod_TempCleaner"
Option Compare Database
Option Explicit

'-------------------------------------
' Module: acc_mod_TempCleaner
' 説明  : VBAコンポーネントの削除（スタンドアロン版）
' 備考  : 共通部品に依存せず、単体で動作します。
'-------------------------------------
'-------------------------------------
' 手順：
' 1. 下の「ExecDelete」という文字の中にカーソルを置く
' 2. キーボードの「F5」キーを1回押す
'-------------------------------------
Public Sub ExecDelete()
    ' クラスモジュールをすべて消したい場合はこのまま
    ' 別の名前にしたい場合は "cls*" の部分だけ書き換えてください
    Call SimpleCleanStandalone("mod_IcubeImport*")
End Sub

' 実際の削除ロジック（ここはいじらなくてOKです）
Private Sub SimpleCleanStandalone(ByVal pattern As String)
    On Error Resume Next
    Dim i As Long
    Dim targetNames As New Collection
    Dim mdl As Object

    ' 1. まず名前をリストアップする
    For Each mdl In CurrentProject.AllModules
        If mdl.Name Like pattern Then
            targetNames.Add mdl.Name
        End If
    Next mdl

    ' 2. リストをもとに削除実行
    If targetNames.count = 0 Then
        MsgBox "対象が見つかりませんでした: " & pattern
        Exit Sub
    End If

    For i = 1 To targetNames.count
        DoCmd.Close acModule, targetNames(i), acSaveNo
        DoCmd.DeleteObject acModule, targetNames(i)
        Debug.Print "削除成功: " & targetNames(i)
    Next i

    MsgBox targetNames.count & " 個のモジュールを削除しました。"
End Sub


'=================================================
' サブルーチン名 : DeleteModuleStandalone
' 引数   : moduleName (String) - 削除対象の名前
' 説明   : 指定されたモジュールまたはクラスを削除する。
'=================================================
Public Sub DeleteModuleStandalone(ByVal moduleName As String)
    On Error GoTo ErrHandler

    ' 自身のモジュールを削除しようとした場合はスキップ（異常終了防止）
    If moduleName = Application.VBE.ActiveCodePane.CodeModule.Name Then
        Debug.Print "Skip: Cannot delete the currently running module [" & moduleName & "]"
        Exit Sub
    End If

    ' オブジェクトを閉じてから削除（編集中のロック回避）
    On Error Resume Next
    DoCmd.Close acModule, moduleName, acSaveNo
    On Error GoTo ErrHandler

    ' 削除実行
    DoCmd.DeleteObject acModule, moduleName
    Debug.Print "Deleted: " & moduleName

    Exit Sub

ErrHandler:
    ' エラー番号 2008: オブジェクトが存在しない場合は無視
    If Err.Number <> 2008 Then
        MsgBox "Error deleting " & moduleName & ": " & Err.description, vbCritical
    End If
End Sub

'=================================================
' サブルーチン名 : CleanModulesByPattern
' 引数   : pattern (String) - 削除条件（例: "tmp_*" / "at_clsTest_*"）
' 説明   : パターンに一致するモジュールを一括削除する。
'=================================================
Public Sub CleanModulesByPattern(ByVal pattern As String)
    Dim moduleNames As New Collection
    Dim i As Long
    Dim currentMod As String
    
    ' 実行中モジュール名を取得
    currentMod = Application.VBE.ActiveCodePane.CodeModule.Name

    ' 1. 削除対象のリストアップ（ループ中の削除によるエラー回避）
    Dim mdl As Object
    For Each mdl In CurrentProject.AllModules
        If mdl.Name Like pattern And mdl.Name <> currentMod Then
            moduleNames.Add mdl.Name
        End If
    Next mdl

    ' 2. 一括削除の実行
    If moduleNames.count = 0 Then
        Debug.Print "No modules found matching pattern: " & pattern
        Exit Sub
    End If

    For i = 1 To moduleNames.count
        DeleteModuleStandalone moduleNames.item(i)
    Next i
    
    MsgBox moduleNames.count & " items deleted.", vbInformation
End Sub

