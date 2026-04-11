VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_DataLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module: cls_DataLookup
' 説明　：テーブルデータのキャッシュ・検索・連結処理を行う汎用クラス
' 作成日：2025/11/14
'-------------------------------------
Option Explicit

' --- メンバ変数 ---
Private pDictData As Object ' Scripting.Dictionary (検索用キャッシュ)
Private pDelimiter As String

' --- プロパティ ---
Public Property Let Delimiter(ByVal value As String)
    pDelimiter = value
End Property

'============================================
' メソッド名 : Initialize
' 概要       : 指定テーブルを読み込み、検索用辞書を作成する
' 引数       : tbl As ListObject - ソーステーブル（tbl_内訳）
'            : keyColName As String - キー列名（内訳ID）
'            : valColName As String - 取得値列名（小分類 or 修繕内容）
'============================================
Public Sub Initialize(tbl As ListObject, keyColName As String, valColName As String)
    
    Dim dataArr As Variant
    Dim keyIdx As Long, valIdx As Long
    Dim i As Long
    Dim key As String, val As String
    
    ' デフォルト区切り文字
    If pDelimiter = "" Then pDelimiter = ","
    
    ' Dictionary初期化（連想配列）
    Set pDictData = CreateObject("Scripting.Dictionary")
    
    ' データ配列取得（高速化）
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    dataArr = tbl.Range.value ' ヘッダー含む全データ
    
    ' 列インデックス特定
    keyIdx = GetColIndex(dataArr, keyColName)
    valIdx = GetColIndex(dataArr, valColName)
    
    If keyIdx = 0 Or valIdx = 0 Then
        Err.Raise 1000, "cls_DataLookup", "指定された列が見つかりません: " & keyColName & ", " & valColName
    End If
    
    ' データを辞書に格納（重複除去しながら連結）
    For i = 2 To UBound(dataArr, 1) ' 1行目はヘッダーなので2行目から
        key = CStr(dataArr(i, keyIdx))
        val = CStr(dataArr(i, valIdx))
        
        If key <> "" And val <> "" Then
            If Not pDictData.Exists(key) Then
                ' 新規キー登録
                pDictData.Add key, val
            Else
                ' 既存キーなら、値が未登録の場合のみ連結（値の重複回避）
                If InStr(1, pDelimiter & pDictData(key) & pDelimiter, pDelimiter & val & pDelimiter) = 0 Then
                    pDictData(key) = pDictData(key) & pDelimiter & val
                End If
            End If
        End If
    Next i
    
End Sub

'============================================
' メソッド名 : GetValue
' 概要       : IDに対応する連結値を取得する
' 引数       : id As String - 検索キー
' 戻り値     : String - 連結された値（見つからない場合は空文字）
'============================================
Public Function GetValue(id As String) As String
    If pDictData Is Nothing Then Exit Function
    If pDictData.Exists(id) Then
        GetValue = pDictData(id)
    Else
        GetValue = ""
    End If
End Function

'============================================
' メソッド名 : ExecuteBatchUpdate
' 概要       : 対象シートの列を一括更新する
' 引数       : targetWs As Worksheet - 対象シート（計画表）
'            : keyColLetter As String - キー列（G）
'            : outColLetter As String - 出力列（H or I）
'            : startRow As Long - 開始行（9 or 15）
'============================================
Public Sub ExecuteBatchUpdate(targetWs As Worksheet, keyColLetter As String, outColLetter As String, startRow As Long)
    
    Dim lastRow As Long
    Dim targetRange As Range
    Dim keys As Variant, results As Variant
    Dim i As Long
    Dim keyVal As String
    
    ' 最終行取得
    lastRow = targetWs.Cells(targetWs.Rows.count, keyColLetter).End(xlUp).Row
    If lastRow < startRow Then Exit Sub
    
    ' 範囲設定
    Set targetRange = targetWs.Range(targetWs.Cells(startRow, keyColLetter), targetWs.Cells(lastRow, keyColLetter))
    
    ' 配列で読み込み（高速化）
    keys = targetRange.value
    ReDim results(1 To UBound(keys, 1), 1 To 1)
    
    ' 一括変換
    For i = 1 To UBound(keys, 1)
        keyVal = Trim(CStr(keys(i, 1)))
        results(i, 1) = Me.GetValue(keyVal)
    Next i
    
    ' 結果書き込み
    targetWs.Cells(startRow, outColLetter).Resize(UBound(results, 1), 1).value = results
    
End Sub

' --- 内部ヘルパー ---
Private Function GetColIndex(arr As Variant, colName As String) As Long
    Dim c As Long
    For c = 1 To UBound(arr, 2)
        ' 両方をTrimしてから比較することで、スペースの有無に関わらず一致させる
        If Trim(CStr(arr(1, c))) = Trim(colName) Then
            GetColIndex = c
            Exit Function
        End If
    Next c
    GetColIndex = 0
End Function


' ==============================================================================
'  以下、cls_DataLookup の既存コードの末尾に追加してください
' ==============================================================================

'============================================
' メソッド名 : ExecuteDataExtraction
' 概要       : 内訳データを抽出し、対象シートの指定行以降に出力する
' 引数       : sourceTbl As ListObject - 内訳テーブル
'            : targetWs As Worksheet - 出力先シート
'            : filterId As String - 抽出するID
'            : startRow As Long - 出力開始行 (例: 13)
'            : titleRow As Long - タイトル行 (例: 7)
'============================================
Public Sub ExecuteDataExtraction(sourceTbl As ListObject, targetWs As Worksheet, filterId As String, startRow As Long, titleRow As Long)
    
    Dim srcData As Variant, srcHeader As Variant
    Dim resData() As Variant
    Dim colMap As Collection
    Dim r As Long, c As Long, outRow As Long
    Dim idColIdx As Long
    Dim srcCol As Long, tgtCol As Long
    Dim matchCount As Long
    
    ' 1. データ読み込み
    If sourceTbl.DataBodyRange Is Nothing Then Exit Sub
    srcData = sourceTbl.DataBodyRange.value
    srcHeader = sourceTbl.HeaderRowRange.value
    
    ' ID列の特定
    idColIdx = GetColIndex(srcHeader, "内訳ID")
    If idColIdx = 0 Then Err.Raise 1002, , "内訳ID列が見つかりません"
    
    ' 2. マッピング作成（シートのタイトル行とテーブルヘッダーを照合）
    Set colMap = New Collection
    For c = 1 To 22 ' A〜V列程度を想定
        Dim tTitle As String
        tTitle = Trim(CStr(targetWs.Cells(titleRow, c).value))
        If tTitle <> "" Then
            srcCol = GetColIndex(srcHeader, tTitle)
            If srcCol > 0 Then
                ' 配列(元列, 先列)を保存
                colMap.Add Array(srcCol, c)
            End If
        End If
    Next c
    
    If colMap.count = 0 Then Exit Sub
    
    ' 3. データ抽出（配列内で処理）
    ' まずサイズ確保（最大行数）
    ReDim resData(1 To UBound(srcData, 1), 1 To 22) ' 22列分確保
    
    outRow = 0
    For r = 1 To UBound(srcData, 1)
        If CStr(srcData(r, idColIdx)) = filterId Then
            outRow = outRow + 1
            ' マッピングに従って列をコピー
            Dim mapItem As Variant
            For Each mapItem In colMap
                resData(outRow, mapItem(1)) = srcData(r, mapItem(0))
            Next mapItem
        End If
    Next r
    
    ' 4. シートへ書き出し
    ' 既存データクリア
    targetWs.Range("A" & startRow & ":V" & (startRow + 1000)).ClearContents
    
    If outRow > 0 Then
        ' データ部分だけ書き出し
        targetWs.Range("A" & startRow).Resize(outRow, 22).value = resData
    End If
    
End Sub


