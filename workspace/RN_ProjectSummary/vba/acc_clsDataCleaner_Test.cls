VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsDataCleaner_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsDataCleaner_Test
' 説明：検証用・実況ログ付きクレンジングクラス
'===========================================================
Option Compare Database
Option Explicit

Private pSanitizer As acc_clsTextSanitizer
Private pDateMath  As com_clsDateMath

Private Sub Class_Initialize()
    On Error Resume Next
    Set pSanitizer = New acc_clsTextSanitizer
    Set pDateMath = New com_clsDateMath
    If pSanitizer Is Nothing Then Debug.Print "!!警告: acc_clsTextSanitizerの生成に失敗しました"
End Sub

'--- 文字列洗浄（ログ付き） ---
Public Function CleanText(ByVal val As Variant, Optional ByVal defaultValue As String = "") As String
    Dim result As String
    If IsNull(val) Then
        result = defaultValue
    Else
        result = Trim(CStr(val))
    End If
    CleanText = result
End Function

'--- 通貨型変換：TextToCurrency（★エラー修正の核心） ---
Public Function TextToCurrency(ByVal val As Variant, Optional ByVal defaultValue As Currency = 0) As Currency
    Dim result As Currency
    If IsNull(val) Or Not IsNumeric(val) Then
        result = defaultValue
    Else
        On Error Resume Next
        result = CCur(val)
        If Err.Number <> 0 Then result = defaultValue
    End If
    ' ログを出す
    Debug.Print "      [Conv] Currency化: [" & val & "] -> " & result
    TextToCurrency = result
End Function

'--- 数値変換：TextToDouble ---
Public Function TextToDouble(ByVal val As Variant, Optional ByVal defaultValue As Double = 0) As Double
    Dim result As Double
    If IsNull(val) Or Not IsNumeric(val) Then
        result = defaultValue
    Else
        On Error Resume Next
        result = CDbl(val)
        If Err.Number <> 0 Then result = defaultValue
    End If
    TextToDouble = result
End Function

'--- 日付変換：TextToDate ---
Public Function TextToDate(ByVal val As Variant, Optional ByVal defaultDate As Date = #1/1/1900#) As Date
    Dim result As Date
    If IsNull(val) Or val = "" Then
        result = defaultDate
    ElseIf IsDate(val) Then
        result = CDate(val)
    Else
        pDateMath.RawValue = CStr(val)
        If pDateMath.IsValid Then
            result = DateSerial(pDateMath.Year, pDateMath.Month, pDateMath.Day)
        Else
            result = defaultDate
        End If
    End If
    Debug.Print "      [Conv] Date化: [" & val & "] -> " & result
    TextToDate = result
End Function
