Attribute VB_Name = "she02_IdMappingDouble"
'-------------------------------------
' Module: she02_IdMappingDouble
' 説明　：内訳IDマッピング処理とIdpa値による上行コピー処理の統合実行モジュール
' 作成日：2025/11/14
' 更新日：-
'-------------------------------------

Option Explicit

'============================================
' Module: she02_IdMappingDouble
' プロシージャ名         : ExecuteIdMappingAndIdpaProcess
' 概要                   : 内訳IDマッピング処理実行後、Idpa値による上行コピー処理を実行する
' 引数                   : なし
' 戻り値（Functionのみ） : なし
' 呼び出し元フォーム／イベント : 手動実行またはボタンイベント
' 関連情報               : she02_IdMappingモジュールの処理を呼び出し後、追加処理を実行
' 備考                   : 2段階の処理を順次実行する統合プロシージャ
'============================================
Public Sub ExecuteIdMappingAndIdpaProcess()
    
    ' --- 1. 内訳IDマッピング処理実行 ---
    Call ExecuteBreakdownIdMappingSingle
    
    ' --- 2. Idpa値による上行コピー処理実行 ---
    Call ProcessIdpaBasedCopy
    
    ' --- 3. 完了メッセージ ---
    MsgBox "全ての処理が完了しました。" & vbCrLf & _
           "1. 内訳IDマッピング処理" & vbCrLf & _
           "2. Idpa値による上行コピー処理", vbInformation, "処理完了"
    
End Sub

'============================================
' プロシージャ名         : ProcessIdpaBasedCopy
' 概要                   : Idpa列が"b"の行の内訳IDを1行上の内訳ID列にコピーする
' 引数                   : なし
' 戻り値（Functionのみ） : なし
' 呼び出し元フォーム／イベント : ExecuteIdMappingAndIdpaProcess
' 関連情報               : テーブル「tbl_内訳」内でのコピー処理
' 備考                   : 1行目はスキップし、2行目以降を処理対象とする
'============================================
Private Sub ProcessIdpaBasedCopy()
    
    ' --- 1. 初期化 ---
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim lastRow As Long
    Dim i As Long
    Dim copyCount As Long
    
    ' --- 2. ワークシート・テーブル取得 ---
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets("内訳")
    Set tbl = ws.ListObjects("tbl_内訳")
    Set dataRange = tbl.DataBodyRange
    
    ' テーブルが空の場合は処理終了
    If dataRange Is Nothing Then
        MsgBox "テーブルにデータがありません。", vbExclamation, "処理終了"
        Exit Sub
    End If
    
    lastRow = dataRange.Rows.count
    
    ' --- 3. 列インデックス取得 ---
    Dim idpaColIndex As Long
    Dim breakdownIdColIndex As Long
    
    idpaColIndex = GetColumnIndex(tbl, "Idpa")
    breakdownIdColIndex = GetColumnIndex(tbl, "内訳ID")
    
    ' 列が見つからない場合はエラー
    If idpaColIndex = 0 Or breakdownIdColIndex = 0 Then
        MsgBox "必要な列が見つかりません。" & vbCrLf & _
               "確認対象：Idpa列、内訳ID列", vbCritical, "エラー"
        Exit Sub
    End If
    
    ' --- 4. メイン処理（2行目から開始） ---
    copyCount = 0
    
    For i = 2 To lastRow
        ' Idpa列の値をチェック
        If CStr(dataRange.Cells(i, idpaColIndex).value) = "b" Then
            ' 現在行の内訳ID値を取得
            Dim currentBreakdownId As String
            currentBreakdownId = CStr(dataRange.Cells(i, breakdownIdColIndex).value)
            
            ' 1行上の内訳ID列に値をコピー
            dataRange.Cells(i - 1, breakdownIdColIndex).value = currentBreakdownId
            
            copyCount = copyCount + 1
        End If
    Next i
    
    ' --- 5. 処理結果表示 ---
    MsgBox "Idpa値による上行コピー処理が完了しました。" & vbCrLf & _
           "処理件数：" & copyCount & "件", vbInformation, "処理完了"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical, "エラー"
    
End Sub

'============================================
' プロシージャ名         : GetColumnIndex
' 概要                   : テーブル内の指定列名の列インデックスを取得する
' 引数                   : tbl As ListObject - 対象テーブル
'                        : columnName As String - 列名
' 戻り値（Functionのみ） : Long - 列インデックス（見つからない場合は0）
' 呼び出し元フォーム／イベント : ProcessIdpaBasedCopy
' 関連情報               : テーブルのヘッダー行から列位置を特定
' 備考                   : 列が見つからない場合は0を返す
'============================================
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    
    Dim i As Long
    
    ' テーブルの各列をチェック
    For i = 1 To tbl.ListColumns.count
        If tbl.ListColumns(i).Name = columnName Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    
    ' 見つからない場合は0を返す
    GetColumnIndex = 0
    
End Function

