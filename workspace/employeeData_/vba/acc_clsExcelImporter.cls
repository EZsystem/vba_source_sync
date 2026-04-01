VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsExcelImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsExcelImporter (EZsystem Refactored)
' 説明　：Excel配列からAccessテーブルへの高速取込を提供するクラス
' 更新日：2026/03/20 - プロパティ名の統一、列名正規化の実装
'=====================================
Option Compare Database
Option Explicit

' --- 内部保持変数 ---
Private pDB               As DAO.Database
Private pDictMapping      As Object ' タイトル -> フィールド名
Private pColIndex         As Object ' タイトル -> 配列列番号
Private pSkipConditions   As Object
Private pDefaultValues    As Object

' --- 公開変数（設定用） ---
Public TargetTable     As String    ' 本番テーブル名（tableNameのエイリアス）
Public TempTable       As String    ' 一時テーブル名（TempTableNameのエイリアス）
Public TitleRow        As Long      ' タイトル行（初期値 1）
Public DataStartRow    As Long      ' データ開始行（初期値 2）
Public StartColumn     As Long      ' 開始列（初期値 1）
Public ImportColumnMax As Long      ' 最大列数（0の場合は自動判定）

' --- 以前のコードとの互換用プロパティ ---
Public Property Let tableName(ByVal val As String): TargetTable = val: End Property
Public Property Let TempTableName(ByVal val As String): TempTable = val: End Property

'=================================================
' サブルーチン名 : Init
' 機能 : 各種辞書オブジェクトとDB接続を初期化する
'=================================================
Public Sub Init()
    Set pDB = CurrentDb
    Set pDictMapping = CreateObject("Scripting.Dictionary")
    Set pColIndex = CreateObject("Scripting.Dictionary")
    Set pSkipConditions = CreateObject("Scripting.Dictionary")
    Set pDefaultValues = CreateObject("Scripting.Dictionary")
    
    ' デフォルト値設定
    TitleRow = 1
    DataStartRow = 2
    StartColumn = 1
End Sub

'=================================================
' プロパティ : MappingMode
' 機能 : "タイトル一致" の場合、自動でマッピングを生成する
'=================================================
Public Property Let MappingMode(ByVal mode As String)
    If mode = "タイトル一致" Then
        Dim key As Variant
        For Each key In pColIndex.Keys
            pDictMapping(key) = key
        Next key
    End If
End Property

'=================================================
' サブルーチン名 : PrepareFieldNames
' 引数 : titleRowArr - タイトル行の2次元配列
' 機能 : タイトルを正規化して列位置を記憶する（表記ゆれ対応）
'=================================================
Public Sub PrepareFieldNames(ByVal titleRowArr As Variant)
    Dim i As Long
    Dim rawTitle As String
    Dim normTitle As String
    
    pColIndex.RemoveAll
    
    ' 配列の列数は UBound(titleRowArr, 2)
    For i = 1 To UBound(titleRowArr, 2)
        rawTitle = Nz(titleRowArr(1, i), "")
        If rawTitle <> "" Then
            ' EZsystem標準：正規化して登録
            normTitle = Normalize_Text(rawTitle)
            If Not pColIndex.Exists(normTitle) Then
                pColIndex.Add normTitle, i
            End If
        End If
    Next i
End Sub

'=================================================
' サブルーチン名 : ImportFromArray
' 引数 : dataArr - 取込対象の純粋なデータ配列
' 機能 : 高速に Recordset へ流し込む
'=================================================
Public Sub ImportFromArray(ByVal dataArr As Variant)
    Dim rs As DAO.Recordset
    Dim r As Long
    Dim titleKey As Variant
    Dim targetField As String
    Dim valCell As Variant
    
    If TempTable = "" Then Err.Raise 999, , "TempTableが指定されていません"
    
    Set rs = pDB.OpenRecordset(TempTable, dbOpenDynaset)
    
    On Error GoTo ErrLine
    For r = 1 To UBound(dataArr, 1)
        If Not ShouldSkip(dataArr, r) Then
            rs.AddNew
            ' マッピング定義されている列のみループ
            For Each titleKey In pDictMapping.Keys
                ' タイトルがExcel列に存在するか確認
                If pColIndex.Exists(titleKey) Then
                    targetField = pDictMapping(titleKey)
                    valCell = dataArr(r, pColIndex(titleKey))
                    
                    ' 空白時のデフォルト値補完
                    If (IsNull(valCell) Or valCell = "") Then
                        If pDefaultValues.Exists(targetField) Then
                            valCell = pDefaultValues(targetField)
                        End If
                    End If
                    
                    rs.Fields(targetField).value = valCell
                End If
            Next titleKey
            rs.Update
        End If
    Next r
    
    rs.Close
    Exit Sub

ErrLine:
    MsgBox "行 " & r & " の取込中にエラーが発生しました: " & Err.Description, vbCritical
    If Not rs Is Nothing Then rs.Close
End Sub

'=================================================
' 関数名 : ShouldSkip (内部用)
'=================================================
Private Function ShouldSkip(dataArr As Variant, r As Long) As Boolean
    ' 必要に応じてスキップロジックを実装（現状は全件通す）
    ShouldSkip = False
End Function

'=================================================
' サブルーチン名 : CommitToMain
' 機能 : 一時テーブルから本番テーブルへ一括追加
'=================================================
Public Sub CommitToMain()
    If TargetTable = "" Or TempTable = "" Then Exit Sub
    Dim sql As String
    sql = "INSERT INTO [" & TargetTable & "] SELECT * FROM [" & TempTable & "];"
    pDB.Execute sql, dbFailOnError
End Sub

'=================================================
' サブルーチン名 : ImportFromWorksheet
' 引数 : ws - Excelのワークシートオブジェクト
'=================================================
Public Sub ImportFromWorksheet(ByVal ws As Object)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim titleArr As Variant
    Dim dataArr As Variant
    
    ' 列数の自動判定
    If ImportColumnMax = 0 Then
        lastCol = ws.Cells(TitleRow, ws.Columns.Count).End(-4159).Column ' xlToLeft
        ImportColumnMax = lastCol - StartColumn + 1
    End If
    
    ' タイトル取得と解析
    titleArr = ws.Range(ws.Cells(TitleRow, StartColumn), ws.Cells(TitleRow, StartColumn + ImportColumnMax - 1)).value
    Call PrepareFieldNames(titleArr)
    
    ' 最終行取得
    lastRow = ws.Cells(ws.rows.Count, StartColumn).End(-4162).Row ' xlUp
    If lastRow < DataStartRow Then Exit Sub ' データなし
    
    ' データ一括取得
    dataArr = ws.Range(ws.Cells(DataStartRow, StartColumn), ws.Cells(lastRow, StartColumn + ImportColumnMax - 1)).value
    
    ' 配列から取込
    Call ImportFromArray(dataArr)
End Sub

'=================================================
' サブルーチン名 : ImportUnpivotedData (New!)
' 引数 : lo - ListObject
'        worksiteName - 作業所名（固定値として追加）
'        sourcePath - 元ファイルパス（固定値として追加）
' 機能 : 横持ちデータを縦持ちに変換しながら高速インポート
'=================================================
Public Sub ImportUnpivotedData( _
    ByVal lo As Object, _
    ByVal worksiteName As String, _
    ByVal sourcePath As String)
    
    Dim rs      As DAO.Recordset
    Dim dataArr As Variant
    Dim r As Long, c As Long
    Dim colName As String
    Dim colNorm As String
    
    ' 固定列のインデックス取得（Normalize_Textを使用）
    Dim idxNo As Long:    idxNo = Get_ColumnIndex_Robust(lo, "No")
    Dim idxYear As Long:  idxYear = Get_ColumnIndex_Robust(lo, "年月")
    Dim idxCode As Long:  idxCode = Get_ColumnIndex_Robust(lo, "工事コード")
    Dim idxName As Long:  idxName = Get_ColumnIndex_Robust(lo, "工事名")
    Dim idxComm As Long:  idxComm = Get_ColumnIndex_Robust(lo, "コメント")
    
    If TempTable = "" Then Err.Raise 999, , "TempTableが指定されていません"
    
    dataArr = lo.DataBodyRange.value
    Set rs = pDB.OpenRecordset(TempTable, dbOpenDynaset)
    
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To lo.ListColumns.Count
            colName = lo.ListColumns(c).Name
            colNorm = Normalize_Text(colName)
            
            ' メタデータ列や空名の列を除外
            If InStr(colName, vbLf) = 0 And _
               UCase(colNorm) <> "NO" And _
               colNorm <> Normalize_Text("年月") And _
               colNorm <> Normalize_Text("工事コード") And _
               colNorm <> Normalize_Text("工事名") And _
               colNorm <> Normalize_Text("コメント") And _
               colName <> "" Then
               
                ' 値がある場合のみレコード作成
                If Not IsEmpty(dataArr(r, c)) And dataArr(r, c) <> 0 Then
                    rs.AddNew
                    
                    ' 固定情報
                    rs!元ファイルパス = sourcePath
                    rs!作業所名 = worksiteName
                    
                    ' 元データの固定マッピング
                    If idxNo > 0 Then rs!No = dataArr(r, idxNo)
                    If idxYear > 0 Then rs!年月 = dataArr(r, idxYear)
                    If idxCode > 0 Then rs!工事コード = dataArr(r, idxCode)
                    If idxName > 0 Then rs!工事名 = dataArr(r, idxName)
                    If idxComm > 0 Then rs!コメント = dataArr(r, idxComm)
                    
                    ' 展開情報（社員名と割合）
                    rs!社員名 = colName
                    rs!兼務率割合 = dataArr(r, c)
                    
                    rs.Update
                End If
            End If
        Next c
    Next r
    
    rs.Close
End Sub
