Attribute VB_Name = "acc_modFieldInfo"
'-------------------------------------
' Module: acc_modFieldInfo
' 説明   : テーブル名指定によるフィールド情報表示ユーティリティ
' 作成日 : 2025/05/23
' 更新日 : -
'-------------------------------------
Option Compare Database
Option Explicit

'=================================================
' サブルーチン名 : Show_FieldInfo_FromList
' 説明   : テーブル一覧から選択したテーブルの
'          フィールド情報をデバッグ出力する
' 引数   : なし
' 戻り値 : なし
'=================================================
Public Sub Show_FieldInfo_FromList()
    ' --- 1. DAO.Database 取得 ---
    Dim Db As DAO.Database
    Set Db = CurrentDb
    
    ' --- 2. テーブル一覧取得 ---
    Dim td As DAO.TableDef
    Dim tableList As String
    For Each td In Db.TableDefs
        If Left(td.Name, 4) <> "MSys" Then
            tableList = tableList & td.Name & ";"
        End If
    Next td
    If tableList = "" Then
        MsgBox "表示可能なテーブルがありません", vbExclamation
        Exit Sub
    End If
    
    ' --- 3. 配列化 ---
    Dim tableArray() As String
    tableArray = Split(Left(tableList, Len(tableList) - 1), ";")
    
    ' --- 4. テーブル選択 ---
    Dim selectedTable As Variant
    selectedTable = Get_ListChoice("テーブルを選択してください", "テーブル選択", tableArray)
    If selectedTable = "" Then
        MsgBox "処理をキャンセルしました", vbInformation
        Exit Sub
    End If
    
    ' --- 5. フィールド情報表示 ---
    Display_FieldInfo CStr(selectedTable)
End Sub ' ← Show_FieldInfo_FromList 終了

'=================================================
' 関数名 : Get_ListChoice
' 説明   : 項目一覧から番号入力により選択肢を返す
' 引数   : prompt（String） プロンプト文言
'        : title（String）   入力ボックスタイトル
'        : itemList（Variant 配列） 選択項目リスト
' 戻り値 : Variant 選択項目または空文字
'=================================================
Private Function Get_ListChoice(prompt As String, title As String, itemList As Variant) As Variant
    ' --- 1. メッセージ組立 ---
    Dim i As Long
    Dim msg As String
    msg = prompt & vbCrLf & vbCrLf
    For i = LBound(itemList) To UBound(itemList)
        msg = msg & (i + 1) & ". " & itemList(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "番号を入力してください : "
    
    ' --- 2. ユーザー入力 & 検証 ---
    Dim resp As Variant
    resp = InputBox(msg, title)
    If IsNumeric(resp) Then
        If resp >= 1 And resp <= UBound(itemList) + 1 Then
            Get_ListChoice = itemList(resp - 1)
        End If
    End If
    ' 空文字の場合はキャンセルとみなす
End Function ' ← Get_ListChoice 終了

'=================================================
' サブルーチン名 : Display_FieldInfo
' 説明   : 指定テーブルのフィールド名と型を
'          デバッグウィンドウに出力する
' 引数   : tableName（String） テーブル名
' 戻り値 : なし
'=================================================
Public Sub Display_FieldInfo(TableName As String)
    ' --- 1. DAO.Database 取得 ---
    Dim Db As DAO.Database
    Set Db = CurrentDb
    
    ' --- 2. テーブル存在確認 ---
    On Error Resume Next
    Dim tdef As DAO.TableDef
    Set tdef = Db.TableDefs(TableName)
    On Error GoTo 0
    If tdef Is Nothing Then
        MsgBox "テーブルが見つかりません : " & TableName, vbExclamation
        Exit Sub
    End If
    
    ' --- 3. フィールド情報出力 ---
    Dim fld As DAO.Field
    Debug.Print "【テーブル名】 " & TableName
    Debug.Print String(30, "-")
    For Each fld In tdef.fields
        Debug.Print fld.Name & " : " & GetFieldTypeName(fld.Type)
    Next fld
    Debug.Print String(30, "-")
    Debug.Print ""
End Sub ' ← Display_FieldInfo 終了

'=================================================
' 関数名 : GetFieldTypeName
' 説明   : DAO.Field.Type 定数から日本語型名を返す
' 引数   : fieldType（Integer） DAO.Field.Type
' 戻り値 : String 型名称
'=================================================
Private Function GetFieldTypeName(fieldType As Integer) As String
    Select Case fieldType
        Case dbBoolean:  GetFieldTypeName = "Yes/No型"
        Case dbByte:     GetFieldTypeName = "バイト型"
        Case dbInteger:  GetFieldTypeName = "整数型"
        Case dbLong:     GetFieldTypeName = "長整数型"
        Case dbSingle:   GetFieldTypeName = "単精度型"
        Case dbDouble:   GetFieldTypeName = "倍精度型"
        Case dbCurrency: GetFieldTypeName = "通貨型"
        Case dbDate:     GetFieldTypeName = "日付/時刻型"
        Case dbText:     GetFieldTypeName = "テキスト型"
        Case dbMemo:     GetFieldTypeName = "メモ型"
        Case dbGUID:     GetFieldTypeName = "GUID型"
        Case Else:       GetFieldTypeName = "不明型(" & fieldType & ")"
    End Select
End Function ' ← GetFieldTypeName 終了

