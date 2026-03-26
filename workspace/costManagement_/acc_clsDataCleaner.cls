VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsDataCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class: acc_clsDataCleaner
' 説明 : データ型変換とクレンジングを提供する汎用クラス
'=====================================
Option Compare Database
Option Explicit

' 文字列のクレンジング（NULLなら空文字を返す）
Public Function CleanText(ByVal val As Variant) As String
    If IsNull(val) Then
        CleanText = ""
    Else
        CleanText = Trim(CStr(val))
    End If
End Function ' ← CleanText 終了

' 数値型への安全変換（NULLならデフォルト値を返す）
Public Function CleanNumber(ByVal val As Variant, Optional ByVal defaultValue As Double = 0) As Double
    If IsNull(val) Or Not IsNumeric(val) Then
        CleanNumber = defaultValue
    Else
        CleanNumber = CDbl(val)
    End If
End Function ' ← CleanNumber 終了

' 日付型への安全変換
Public Function CleanDate(ByVal val As Variant, Optional ByVal defaultDate As Date = #1/1/1900#) As Date
    If IsNull(val) Or Not IsDate(val) Then
        CleanDate = defaultDate
    Else
        CleanDate = CDate(val)
    End If
End Function ' ← CleanDate 終了

' テキストから日付型への変換
Public Function TextToDate(ByVal val As Variant, Optional ByVal defaultDate As Date = #1/1/1900#) As Date
    If IsNull(val) Or Trim(CStr(val)) = "" Then
        TextToDate = defaultDate
    ElseIf IsDate(val) Then
        TextToDate = CDate(val)
    Else
        TextToDate = defaultDate
    End If
End Function ' ← TextToDate 終了

' テキストから通貨型への変換
Public Function TextToCurrency(ByVal val As Variant, Optional ByVal defaultValue As Currency = 0) As Currency
    If IsNull(val) Or Trim(CStr(val)) = "" Then
        TextToCurrency = defaultValue
    ElseIf IsNumeric(val) Then
        TextToCurrency = CCur(val)
    Else
        TextToCurrency = defaultValue
    End If
End Function ' ← TextToCurrency 終了

' テキストから倍精度浮動小数点型への変換
Public Function TextToDouble(ByVal val As Variant, Optional ByVal defaultValue As Double = 0) As Double
    If IsNull(val) Or Trim(CStr(val)) = "" Then
        TextToDouble = defaultValue
    ElseIf IsNumeric(val) Then
        TextToDouble = CDbl(val)
    Else
        TextToDouble = defaultValue
    End If
End Function ' ← TextToDouble 終了

' 日本語の型名をVBA型名に変換する
Public Function GetSupportedVBAType(ByVal jpType As String) As String
    Dim typeName As String
    typeName = Trim(jpType)

    Select Case typeName
        Case "テキスト型":   GetSupportedVBAType = "String"
        Case "長整数型":     GetSupportedVBAType = "Long"
        Case "整数型":       GetSupportedVBAType = "Integer"
        Case "倍精度浮動小数点型": GetSupportedVBAType = "Double"
        Case "通貨型":       GetSupportedVBAType = "Currency"
        Case "日付/時刻型":   GetSupportedVBAType = "Date"
        Case "Yes/No型":     GetSupportedVBAType = "Boolean"
        Case Else:           GetSupportedVBAType = "Variant"
    End Select
End Function ' ← GetSupportedVBAType 終了
