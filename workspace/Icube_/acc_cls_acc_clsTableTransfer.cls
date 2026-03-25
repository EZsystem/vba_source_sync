VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTableTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pConn As DAO.Database

Public Sub Init(ByRef db As DAO.Database)
    Set pConn = db
End Sub

'===========================================================
' メソッド : ExecuteTransfer
'===========================================================
Public Sub ExecuteTransfer(ByVal srcTable As String, ByVal tgtTable As String, Optional ByVal keyField As String = "")
    On Error GoTo Err_Handler
    
    ' 1. 型変換を考慮したセレクト用フィールドリストを取得
    Dim selectFields As String
    Dim insertFields As String
    
    ' 専用のヘルパー関数で SQL の部品を作る
    GetFieldLists srcTable, tgtTable, insertFields, selectFields
    
    If insertFields = "" Then Err.Raise 1004, , "共通列がないにゃ。"

    ' 2. 既存レコードの削除 (型不一致を防ぐため Val関数を使用)
    Dim sqlDel As String
    If keyField <> "" Then
        ' 数値型の比較でエラーが出ないよう Val() でテキストを数値化
        sqlDel = "DELETE FROM [" & tgtTable & "] WHERE [" & keyField & "] IN (SELECT Val([" & keyField & "]) FROM [" & srcTable & "]);"
    Else
        sqlDel = "DELETE FROM [" & tgtTable & "];"
    End If
    
    pConn.Execute sqlDel, 128 ' dbFailOnError

    ' 3. データの一括挿入 (空文字を Null に変換する SQL を実行)
    Dim sqlIns As String
    sqlIns = "INSERT INTO [" & tgtTable & "] (" & insertFields & ") SELECT " & selectFields & " FROM [" & srcTable & "];"
    
    ' 生成されたSQLをデバッグ出力（確認用）
    Debug.Print "生成されたSQL: " & Left(sqlIns, 200) & "..."
    
    pConn.Execute sqlIns, 128
    Exit Sub

Err_Handler:
    Err.Raise Err.Number, "acc_clsTableTransfer", Err.description
End Sub

' --- フィールドリスト生成 (空文字を Null に置換するロジック付き) ---
Private Sub GetFieldLists(ByVal src As String, ByVal tgt As String, ByRef outInsert As String, ByRef outSelect As String)
    Dim fldsSrc As DAO.Fields: Set fldsSrc = pConn.TableDefs(src).Fields
    Dim fldsTgt As DAO.Fields: Set fldsTgt = pConn.TableDefs(tgt).Fields
    Dim fSrc As DAO.Field, fTgt As DAO.Field
    Dim insList As String: insList = ""
    Dim selList As String: selList = ""
    
    For Each fSrc In fldsSrc
        ' オートナンバー型(16)は除外
        If (fSrc.Attributes And 16) = 0 Then
            For Each fTgt In fldsTgt
                If StrComp(fSrc.Name, fTgt.Name, 1) = 0 Then
                    ' 挿入先リスト
                    insList = insList & "[" & fTgt.Name & "], "
                    
                    ' 取得元リスト: 空文字("") の場合は Null に変換するように SQL を組む
                    ' これにより、数値型や日付型の列にテキスト型の空文字が入ってエラーになるのを防ぐ
                    selList = selList & "IIf([" & fSrc.Name & "]='', Null, [" & fSrc.Name & "]), "
                    Exit For
                End If
            Next fTgt
        End If
    Next fSrc
    
    If insList <> "" Then
        outInsert = Left(insList, Len(insList) - 2)
        outSelect = Left(selList, Len(selList) - 2)
    End If
End Sub


' フィールド存在チェック
Private Function FieldExists(ByVal tblName As String, ByVal fldName As String) As Boolean
    On Error Resume Next
    Dim f As Object: Set f = pConn.TableDefs(tblName).Fields(fldName)
    FieldExists = (Err.Number = 0)
    On Error GoTo 0
End Function

