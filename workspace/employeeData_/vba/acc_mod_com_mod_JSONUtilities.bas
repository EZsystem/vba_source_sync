Attribute VB_Name = "com_mod_JSONUtilities"
'-------------------------------------
' Module: com_mod_JSONUtilities
' 説明  : JSON シリアライズ／デシリアライズ
' 作成日: 2025/05/04
' 更新日: -
'-------------------------------------
Option Compare Database
Option Explicit

'-------------------------------------
' Module: com_mod_JSONUtilities (EZsystem Refactored)
' 説明  : JSON シリアライズ／デシリアライズの統一インターフェース
' 依存  : JsonConverter.bas (VBA-JSON)
' 参照  : Microsoft Scripting Runtime
'-------------------------------------


'=================================================
' 関数名 : ToJSON
' 機能   : 辞書やコレクションを JSON 文字列に変換
'=================================================
Public Function ToJSON(ByVal value As Variant) As String
    On Error GoTo ErrLine
    ' JsonConverter モジュールが存在することを確認してください
    ToJSON = JsonConverter.ConvertToJson(value)
    Exit Function
ErrLine:
    ToJSON = ""
    Debug.Print "ToJSON Error: " & Err.Description
End Function

'=================================================
' 関数名 : FromJSON / ParseJSON (互換用)
' 機能   : JSON 文字列を Variant（Dictionary/Collection）に変換
'=================================================
Public Function FromJSON(ByVal jsonText As String) As Variant
    Dim result As Variant
    
    On Error GoTo ErrLine
    ' 解析実行
    result = JsonConverter.ParseJson(jsonText)
    
    ' 戻り値がオブジェクト（{} or []）なら Set、それ以外ならそのまま代入
    If IsObject(result) Then
        Set FromJSON = result
    Else
        FromJSON = result
    End If
    Exit Function

ErrLine:
    FromJSON = Nothing
    Debug.Print "FromJSON Error: " & Err.Description
End Function

' 以前の提案コードとの互換性のための別名
Public Function ParseJson(ByVal jsonText As String) As Variant
    If IsObject(FromJSON(jsonText)) Then
        Set ParseJson = FromJSON(jsonText)
    Else
        ParseJson = FromJSON(jsonText)
    End If
End Function
