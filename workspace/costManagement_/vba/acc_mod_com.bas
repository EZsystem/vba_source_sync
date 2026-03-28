Attribute VB_Name = "acc_mod_com"
'-------------------------------------
' Module: acc_mod_com
' 説明   : アプリケーション全体で利用する共通関数群
'-------------------------------------
Option Compare Database
Option Explicit ' 全モジュール必須宣言

'--------------------------------------------
' 関数名 : GetSelectedExcelFiles
' 概要   : ユーザーが選択したExcelファイルのパス一覧をCollectionで返す
'--------------------------------------------
Public Function GetSelectedExcelFiles() As Collection
    Dim fd As FileDialog
    Dim colFiles As New Collection
    Dim varItem As Variant
    
    On Error GoTo ErrHandler
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
         .title = "インポートするExcelファイルを選択してください"
         .AllowMultiSelect = True
         .Filters.Clear
         .Filters.Add "Excelファイル", "*.xls; *.xlsx; *.xlsm", 1
         If .Show = -1 Then
              For Each varItem In .SelectedItems
                   colFiles.Add varItem
              Next varItem
         End If
    End With
    Set GetSelectedExcelFiles = colFiles
    Exit Function
    
ErrHandler:
    MsgBox "ファイル選択ダイアログエラー " & Err.Number & "：" & Err.Description, vbCritical
    Set GetSelectedExcelFiles = Nothing
End Function ' ← GetSelectedExcelFiles 終了



'--------------------------------------------
' 関数名 : FieldExists
' 概要   : 指定したレコードセット内にフィールドが存在するか確認する
'--------------------------------------------
Public Function FieldExists(ByRef rs As DAO.Recordset, ByVal fieldName As String) As Boolean
    Dim i As Integer
    On Error Resume Next
    i = rs.Fields(fieldName).OrdinalPosition
    FieldExists = (Err.Number = 0)
    Err.Clear
End Function ' ← FieldExists 終了

'--------------------------------------------
' 関数名 : ConvertValue
' 概要   : データ型と補完ルールに応じて型変換＋空欄対応を行う
'--------------------------------------------
Public Function ConvertValue(ByVal val As Variant, ByVal datType As Variant, ByVal emptyVal As Variant, ByRef cleaner As acc_clsDataCleaner) As Variant
    Dim typeStr As String
    Dim isVacant As Boolean
    
    ' --- 1. 空欄判定（Null対策を強化） ---
    isVacant = False
    If IsNull(val) Then
        isVacant = True
    ElseIf Trim(CStr(Nz(val, ""))) = "" Then
        isVacant = True
    End If
    
    ' 空欄の場合はデフォルト値を採用
    If isVacant Then val = emptyVal

    ' --- 2. 型変換（datTypeがNullでもエラーにしない） ---
    typeStr = Trim(CStr(Nz(datType, "")))

    Select Case typeStr
        Case "テキスト型": ConvertValue = cleaner.CleanText(val)
        Case "通貨型":     ConvertValue = cleaner.TextToCurrency(val)
        Case "日付型":     ConvertValue = cleaner.TextToDate(val)
        Case "倍精度浮動小数点型": ConvertValue = cleaner.TextToDouble(val)
        Case Else:         ConvertValue = val
    End Select
End Function ' ← ConvertValue 終了
