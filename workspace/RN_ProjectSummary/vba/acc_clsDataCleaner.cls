VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsDataCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "acc_clsDataCleaner"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'===========================================================

' Class Module: acc_clsDataCleaner

' 説明：あらゆるデータ型の洗浄・変換を一括管理する。

'       acc_clsTextSanitizer を内部に持ち、文字列・数値・日付の安全なクレンジングを行う。

' 統合対応：原価管理（costManagement_）プロジェクトで必要な通貨型・浮動小数点型を追加。

' 更新日：2026/03/26（RN/原価管理統合版）

'===========================================================

Option Explicit



' --- 内部保持変数 ---
Private pSanitizer As acc_clsTextSanitizer
Private pDateMath  As com_clsDateMath

'===========================================================
' サブルーチン名 : Class_Initialize
' 機能 : クラスの生成時に内部ユーティリティを自動初期化する。
'===========================================================
Private Sub Class_Initialize()
    Set pSanitizer = New acc_clsTextSanitizer
    Set pDateMath = New com_clsDateMath
End Sub

'===========================================================
' サブルーチン名 : Init
' 機能 : 内部ユーティリティクラスを初期化する（互換用）。
'===========================================================
Public Sub Init()
    If pSanitizer Is Nothing Then Set pSanitizer = New acc_clsTextSanitizer
    If pDateMath Is Nothing Then Set pDateMath = New com_clsDateMath
End Sub



'===========================================================

' 関数名 : CleanText

' 引数 : val - 洗浄対象の値

'        defaultValue - Null等の場合の戻り値

' 戻り値 : String

' 機能 : 文字列化、Trim、改行除去を行い、安全な文字列を返す。

'===========================================================

Public Function CleanText(ByVal val As Variant, Optional ByVal defaultValue As String = "") As String

    CleanText = pSanitizer.CleanText(val, defaultValue)

End Function



'===========================================================

' 関数名 : CleanLong

' 引数 : val - 変換対象の値

'        defaultValue - エラー時の戻り値 (既定: 0)

' 戻り値 : Long

'===========================================================

Public Function CleanLong(ByVal val As Variant, Optional ByVal defaultValue As Long = 0) As Long

    If IsNull(val) Or Not IsNumeric(val) Then

        CleanLong = defaultValue

    Else

        On Error Resume Next

        CleanLong = CLng(val)

        If Err.Number <> 0 Then CleanLong = defaultValue

    End If

End Function



'===========================================================

' 関数名 : CleanDouble

' 引数 : val - 変換対象の値

'        defaultValue - エラー時の戻り値 (既定: 0)

' 戻り値 : Double

'===========================================================

Public Function CleanDouble(ByVal val As Variant, Optional ByVal defaultValue As Double = 0) As Double

    If IsNull(val) Or Not IsNumeric(val) Then

        CleanDouble = defaultValue

    Else

        On Error Resume Next

        CleanDouble = CDbl(val)

        If Err.Number <> 0 Then CleanDouble = defaultValue

    End If

End Function



'===========================================================

' 関数名 : CleanCurrency

' 引数 : val - 変換対象の値

'        defaultValue - エラー時の戻り値 (既定: 0)

' 戻り値 : Currency

' 機能 : 通貨型への変換。原価管理プロジェクトの金額項目等で使用。

'===========================================================

Public Function CleanCurrency(ByVal val As Variant, Optional ByVal defaultValue As Currency = 0) As Currency

    If IsNull(val) Or Not IsNumeric(val) Then

        CleanCurrency = defaultValue

    Else

        On Error Resume Next

        CleanCurrency = CCur(val)

        If Err.Number <> 0 Then CleanCurrency = defaultValue

    End If

End Function



'===========================================================

' 関数名 : CleanDate

' 引数 : val - 変換対象の値

'        defaultDate - 不正な場合の戻り値（既定: 1900/01/01）

' 戻り値 : Date

' 機能 : 8桁/6桁の数値形式を含め、安全に日付型へ変換する。

'===========================================================

Public Function CleanDate(ByVal val As Variant, Optional ByVal defaultDate As Date = #1/1/1900#) As Date

    If IsNull(val) Or val = "" Then

        CleanDate = defaultDate

        Exit Function

    End If

    

    If IsDate(val) Then

        CleanDate = CDate(val)

    Else

        ' 8桁/6桁数字形式の解析を試行 (com_clsDateMathを利用)

        pDateMath.RawValue = CStr(val)

        If pDateMath.IsValid Then

            CleanDate = DateSerial(pDateMath.Year, pDateMath.Month, pDateMath.Day)

        Else

            CleanDate = defaultDate

        End If

    End If

End Function



'===========================================================

' 関数名 : GetSupportedVBAType

' 引数 : jpType - 日本語型名（例：「テキスト型」「長整数型」）

' 戻り値 : VBA型名

' 機能 : Accessの日本語型名をVBAの標準型名にマッピングする。

'===========================================================

'Public Function GetSupportedVBAType(ByVal jpType As String) As String
'
 '   Dim t As String: t = Trim(jpType)

  '  End Select

'End Function

'===========================================================
' --- 以下、acc_clsDataCleaner_Test から統合された新ロジック ---
'===========================================================

'----------------------------------------------------------------
' 関数名 : TextToCurrency
' 概要   : 通貨型への変換。Debug.Print によるログ出力付き。
'----------------------------------------------------------------
Public Function TextToCurrency(ByVal val As Variant, Optional ByVal defaultValue As Currency = 0) As Currency
    Dim result As Currency
    If IsNull(val) Or Not IsNumeric(val) Then
        result = defaultValue
    Else
        On Error Resume Next
        result = CCur(val)
        If Err.Number <> 0 Then result = defaultValue
    End If
    ' イミディエイトウィンドウへログを出力
    Debug.Print "      [Cleanup] Currency化: [" & val & "] -> " & result
    TextToCurrency = result
End Function

'----------------------------------------------------------------
' 関数名 : TextToDouble
' 概要   : 数値（Double）への変換。
'----------------------------------------------------------------
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

'----------------------------------------------------------------
' 関数名 : TextToDate
' 概要   : 日付への変換。Debug.Print によるログ出力付き。
'----------------------------------------------------------------
Public Function TextToDate(ByVal val As Variant, Optional ByVal defaultDate As Date = #1/1/1900#) As Date
    Dim result As Date
    If IsNull(val) Or val = "" Then
        result = defaultDate
    ElseIf IsDate(val) Then
        result = CDate(val)
    Else
        ' 8桁/6桁数字形式の解析を試行
        pDateMath.RawValue = CStr(val)
        If pDateMath.IsValid Then
            result = DateSerial(pDateMath.Year, pDateMath.Month, pDateMath.Day)
        Else
            result = defaultDate
        End If
    End If
    Debug.Print "      [Cleanup] Date化: [" & val & "] -> " & result
    TextToDate = result
End Function

