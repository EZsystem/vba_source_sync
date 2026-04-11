Attribute VB_Name = "she02_IdMappingSingle"
'-------------------------------------
' Module: she02_IdMapping
' 説明　：内訳IDマッピング処理の実行モジュール
' 作成日：2025/10/12
' 更新日：2025/10/17
'-------------------------------------

Option Explicit

'============================================
' Module: she02_IdMapping
' プロシージャ名         : Sub ExecuteBreakdownIdMappingSingle
' 概要                   : 内訳シートと分類シートの連結値を照合し内訳IDをマッピングする
' 引数                   : なし
' 戻り値（Functionのみ） : -
' 呼び出し元フォーム／イベント : 手動実行（将来的にボタンクリック）
' 関連情報               : 大分類・中分類・種類・周期・更新周期を連結して照合
' 備考                   : 一致しない場合は「該当無し」を設定
'============================================
Public Sub ExecuteBreakdownIdMappingSingle()
    
    ' --- 1. 初期化 ---
    Dim wsBreakdown As Worksheet    ' 内訳シート
    Dim wsCategory As Worksheet     ' 分類シート
    Dim tblBreakdown As ListObject  ' 内訳テーブル
    Dim tblCategory As ListObject   ' 分類テーブル
    
    ' --- 2. ワークシート・テーブル取得 ---
    On Error GoTo ErrorHandler
    
    Set wsBreakdown = ThisWorkbook.Worksheets("内訳")
    Set wsCategory = ThisWorkbook.Worksheets("分類")
    Set tblBreakdown = wsBreakdown.ListObjects("tbl_内訳")
    Set tblCategory = wsCategory.ListObjects("tbl_内訳ID")
    
    ' --- 3. データ処理実行 ---
    Call ProcessBreakdownIdMapping(tblBreakdown, tblCategory)
    
    ' --- 4. 完了メッセージ ---
    MsgBox "内訳IDマッピング処理が完了しました", vbInformation, "処理完了"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
    
End Sub   ' ← Subの終わり

'============================================
' プロシージャ名         : Sub ProcessBreakdownIdMapping
' 概要                   : 内訳IDマッピングのメイン処理
' 引数                   : tblSource As ListObject - 記入先テーブル（内訳）
'                        : tblReference As ListObject - 調査先テーブル（分類）
' 戻り値（Functionのみ） : -
' 呼び出し元フォーム／イベント : ExecuteBreakdownIdMapping
' 関連情報               : 連結キーでの照合とIDマッピング
' 備考                   : 大分類・中分類・種類・周期・更新周期を単純連結で照合
'============================================
Private Sub ProcessBreakdownIdMapping(tblSource As ListObject, tblReference As ListObject)
    
    ' --- 1. 変数宣言 ---
    Dim sourceRow As Long           ' 記入先行カウンタ
    Dim referenceRow As Long        ' 調査先行カウンタ
    Dim sourceKey As String         ' 記入先連結キー
    Dim referenceKey As String      ' 調査先連結キー
    Dim foundId As String           ' 発見された内訳ID
    Dim isFound As Boolean          ' 一致フラグ
    
    ' --- 2. 列インデックス取得（記入先：内訳シート） ---
    Dim srcColLargeCategory As Long     ' 大分類列
    Dim srcColMediumCategory As Long    ' 中分類列
    Dim srcColType As Long              ' 種類列
    Dim srcColCycle As Long             ' 周期列
    Dim srcColUpdateCycle As Long       ' 更新周期列
    Dim srcColBreakdownId As Long       ' 内訳ID列
    
    srcColLargeCategory = GetColumnIndex(tblSource, "大分類")
    srcColMediumCategory = GetColumnIndex(tblSource, "中分類")
    srcColType = GetColumnIndex(tblSource, "種類")
    srcColCycle = GetColumnIndex(tblSource, "周期")
    srcColUpdateCycle = GetColumnIndex(tblSource, "更新周期")
    srcColBreakdownId = GetColumnIndex(tblSource, "内訳ID")
    
    ' --- 3. 列インデックス取得（調査先：分類シート） ---
    Dim refColLargeCategory As Long     ' 大分類列
    Dim refColMediumCategory As Long    ' 中分類列
    Dim refColType As Long              ' 種類列
    Dim refColCycle As Long             ' 周期列
    Dim refColUpdateCycle As Long       ' 更新周期列
    Dim refColBreakdownId As Long       ' 内訳ID列
    
    refColLargeCategory = GetColumnIndex(tblReference, "大分類")
    refColMediumCategory = GetColumnIndex(tblReference, "中分類")
    refColType = GetColumnIndex(tblReference, "種類")
    refColCycle = GetColumnIndex(tblReference, "周期")
    refColUpdateCycle = GetColumnIndex(tblReference, "更新周期")
    refColBreakdownId = GetColumnIndex(tblReference, "内訳ID")
    
    ' --- 4. メイン処理ループ ---
    For sourceRow = 1 To tblSource.DataBodyRange.Rows.count
        
        ' --- 4-1. 記入先連結キー作成 ---
        sourceKey = CreateConcatenatedKey(tblSource, sourceRow, _
                                        srcColLargeCategory, srcColMediumCategory, _
                                        srcColType, srcColCycle, srcColUpdateCycle)
        
        ' --- 4-2. 調査先での照合 ---
        isFound = False
        foundId = "該当無し"
        
        For referenceRow = 1 To tblReference.DataBodyRange.Rows.count
            
            ' --- 4-3. 調査先連結キー作成 ---
            referenceKey = CreateConcatenatedKey(tblReference, referenceRow, _
                                                refColLargeCategory, refColMediumCategory, _
                                                refColType, refColCycle, refColUpdateCycle)
            
            ' --- 4-4. キー照合 ---
            If sourceKey = referenceKey Then
                foundId = tblReference.DataBodyRange.Cells(referenceRow, refColBreakdownId).value
                isFound = True
                Exit For
            End If
            
        Next referenceRow
        
        ' --- 4-5. 結果書き込み ---
        tblSource.DataBodyRange.Cells(sourceRow, srcColBreakdownId).value = foundId
        
    Next sourceRow
    
End Sub   ' ← Subの終わり

'============================================
' プロシージャ名         : Function GetColumnIndex
' 概要                   : テーブル内の列名から列インデックスを取得する
' 引数                   : tbl As ListObject - 対象テーブル
'                        : columnName As String - 列名
' 戻り値（Functionのみ） : Long - 列インデックス（1ベース）
' 呼び出し元フォーム／イベント : ProcessBreakdownIdMapping
' 関連情報               : テーブルヘッダーから列位置を特定
' 備考                   : 列が見つからない場合は0を返す
'============================================
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    
    Dim colIndex As Long
    
    ' --- 列名検索 ---
    For colIndex = 1 To tbl.HeaderRowRange.Columns.count
        If tbl.HeaderRowRange.Cells(1, colIndex).value = columnName Then
            GetColumnIndex = colIndex
            Exit Function
        End If
    Next colIndex
    
    ' --- 列が見つからない場合 ---
    GetColumnIndex = 0
    
End Function   ' ← 関数の終わり

'============================================
' プロシージャ名         : Function CreateConcatenatedKey
' 概要                   : 指定された列の値を連結してキーを作成する
' 引数                   : tbl As ListObject - 対象テーブル
'                        : rowIndex As Long - 行インデックス
'                        : col1 As Long, col2 As Long, col3 As Long, col4 As Long, col5 As Long - 列インデックス
' 戻り値（Functionのみ） : String - 連結されたキー
' 呼び出し元フォーム／イベント : ProcessBreakdownIdMapping
' 関連情報               : 大分類・中分類・種類・周期・更新周期を単純連結
' 備考                   : 空白値も含めて連結する
'============================================
Private Function CreateConcatenatedKey(tbl As ListObject, rowIndex As Long, _
                                      col1 As Long, col2 As Long, _
                                      col3 As Long, col4 As Long, _
                                      col5 As Long) As String
    
    Dim key1 As String, key2 As String, key3 As String, key4 As String, key5 As String
    
    ' --- 各列の値を取得 ---
    key1 = CStr(tbl.DataBodyRange.Cells(rowIndex, col1).value)
    key2 = CStr(tbl.DataBodyRange.Cells(rowIndex, col2).value)
    key3 = CStr(tbl.DataBodyRange.Cells(rowIndex, col3).value)
    key4 = CStr(tbl.DataBodyRange.Cells(rowIndex, col4).value)
    key5 = CStr(tbl.DataBodyRange.Cells(rowIndex, col5).value)
    
    ' --- 単純連結 ---
    CreateConcatenatedKey = key1 & key2 & key3 & key4 & key5
    
End Function   ' ← 関数の終わり

