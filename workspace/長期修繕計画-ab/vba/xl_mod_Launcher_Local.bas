Attribute VB_Name = "xl_mod_Launcher_Local"
Option Explicit

' ==========================================
' モジュール名：xl_mod_Launcher_Local
' 説明：ランチャーの起動およびマクロの動的実行制御
' ==========================================

''' <summary>
''' ランチャーフォームを起動する（ルール8-1準拠）
''' </summary>
Public Sub Launcher_Show_Local()
    ' 初期化（ルール4-1）
    ' モーダル（vbModeless）で開くことで、シート操作を可能にする
    frm_Launcher.Show vbModeless
End Sub

''' <summary>
''' 指定されたマクロ名を安全に実行する
''' </summary>
''' <param name="macroName">実行対象のプロシージャ名（フルパス推奨）</param>
Public Sub Launcher_Execute_Macro(ByVal macroName As String)
    ' 1. 初期化・検証（ルール4-1, 11）
    On Error GoTo ErrorHandler
    
    If macroName = "" Then
        MsgBox "マクロ名が指定されていません。", vbExclamation
        Exit Sub
    End If

    ' 2. 実行（XML実装ルール2）
    ' Application.Run を使用し、フォームとの疎結合を維持
    Application.Run macroName

    ' 3. 終了処理
    Exit Sub

ErrorHandler:
    ' エラー出力（ルール4-1, 10）
    MsgBox "マクロの実行中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "実行失敗: " & macroName
End Sub

' ← プロシージャの終わり

