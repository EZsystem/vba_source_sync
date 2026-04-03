Attribute VB_Name = "acc_mod_Import_Json"
Option Explicit

'----------------------------------------------------------------
' Module: acc_mod_Import_Json
' 説明   : VBA-JSON (JsonConverter) を使用した社員情報のインポート
'----------------------------------------------------------------

Private Const TARGET_TABLE As String = "at_testTemp"

'----------------------------------------------------------------
' 関数名 : ImportJsonProcess
' 概要   : JSONファイルを解析し、テーブルへ取り込む
'----------------------------------------------------------------
Public Function ImportJsonProcess(Optional ByVal callingID As Long = 0)
    Dim db          As DAO.Database: Set db = CurrentDb
    Dim rsConfig    As DAO.Recordset
    Dim fso         As Object
    Dim ts          As Object
    Dim rawText     As String
    Dim json        As Object
    Dim item        As Object
    Dim filePath    As String
    Dim recordCount As Long
    
    On Error GoTo Err_Handler
    
    ' 1. レジストリからパス取得 (IDまたは名称)
    Dim strSQL As String
    If callingID > 0 Then
        strSQL = "SELECT [既定パス] FROM [_at_SystemRegistry] WHERE [ID] = " & callingID
    Else
        strSQL = "SELECT [既定パス] FROM [_at_SystemRegistry] WHERE [処理名称] = '社員情報JSONインポート'"
    End If
    
    Set rsConfig = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rsConfig.EOF Then
        MsgBox "JSONインポートの設定がレジストリに見つかりません。" & vbCrLf & _
               "ID: " & callingID, vbCritical
        Exit Function
    End If
    filePath = Nz(rsConfig![既定パス], ""): rsConfig.Close
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        MsgBox "JSONファイルが見つかりません: " & filePath, vbCritical
        Exit Function
    End If

    ' 2. JSON読み込み
    Set ts = fso.OpenTextFile(filePath, 1, False)
    rawText = ts.ReadAll
    ts.Close

    ' 3. JsonConverterによる解析
    ' ※ 参照設定または acc_mod_JsonConverter.bas が必要
    Set json = JsonConverter.ParseJson(rawText)
    
    ' 4. テーブルクリア
    db.Execute "DELETE * FROM " & TARGET_TABLE & ";", dbFailOnError

    ' 5. インポート実行 (トランザクション)
    DBEngine.BeginTrans
    
    ' json が Collection (Array) の場合を想定
    For Each item In json
        Dim sql As String
        ' item(Key) で値を取得。存在しないキーはエラーにならず Null/Empty を返します
        sql = "INSERT INTO " & TARGET_TABLE & " (社員番号, 氏名_戸籍上, 氏名カナ, 氏名_ﾒｰﾙ表示用, 資格, 所属, 役職, 対外呼称) " & _
              "VALUES ('" & item("社員番号") & "', '" & _
                            item("氏名_戸籍上") & "', '" & _
                            item("氏名カナ") & "', '" & _
                            item("氏名_ﾒｰﾙ表示用") & "', '" & _
                            item("資格") & "', '" & _
                            item("所属") & "', '" & _
                            item("役職") & "', '" & _
                            item("対外呼称") & "');"
        db.Execute sql, dbFailOnError
        recordCount = recordCount + 1
    Next item

    DBEngine.CommitTrans
    
    MsgBox "JSONインポート処理が完了しました。件数: " & recordCount & " 件", vbInformation
    Exit Function

Err_Handler:
    On Error Resume Next
    DBEngine.Rollback
    MsgBox "JSONインポート中にエラーが発生しました:" & vbCrLf & Err.Description, vbCritical
End Function
