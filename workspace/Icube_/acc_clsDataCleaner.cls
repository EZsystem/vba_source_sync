VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsDataCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsDataCleaner
' 説明：あらゆるデータ型の洗浄・変換を一括管理する。
'       acc_clsTextSanitizer を内部に持ち、文字列・数値・日付の安全なクレンジングを行う。
' 更新日：2026/03/24（MyLauncher統合・高機能クレンジング版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pSanitizer As acc_clsTextSanitizer
Private pDateMath  As com_clsDateMath

'===========================================================
' サブルーチン名 : Init
' 機能 : 内部ユーティリティクラスを初期化する。
'===========================================================
Public Sub Init()
    Set pSanitizer = New acc_clsTextSanitizer
    Set pDateMath = New com_clsDateMath
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
' 関数名 : CleanDate
' 引数 : val - 変換対象の値
' 戻り値 : Date (不正な場合は 1900/01/01 を返す)
'===========================================================
Public Function CleanDate(ByVal val As Variant) As Date
    Dim minDate As Date: minDate = #1/1/1900#
    
    If IsNull(val) Or val = "" Then
        CleanDate = minDate
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
            CleanDate = minDate
        End If
    End If
End Function

'===========================================================
' 関数名 : GetSupportedVBAType
' 引数 : jpType - 日本語型名（例：「テキスト型」「長整数型」）
' 戻り値 : VBA型名
' 機能 : Accessの日本語型名をVBAの標準型名にマッピングする。
'===========================================================
Public Function GetSupportedVBAType(ByVal jpType As String) As String
    Dim t As String: t = Trim(jpType)
    Select Case t
        Case "テキスト型", "メモ型": GetSupportedVBAType = "String"
        Case "長整数型", "オートナンバー型": GetSupportedVBAType = "Long"
        Case "整数型": GetSupportedVBAType = "Integer"
        Case "倍精度浮動小数点型", "数値型": GetSupportedVBAType = "Double"
        Case "通貨型": GetSupportedVBAType = "Currency"
        Case "日付/時刻型": GetSupportedVBAType = "Date"
        Case "Yes/No型": GetSupportedVBAType = "Boolean"
        Case Else: GetSupportedVBAType = "Variant"
    End Select
End Function
