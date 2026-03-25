VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "com_clsDateMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: com_clsDateMath
' 説明：数字のみの日付文字列（YYYYMMDD / YYYYMM）を解析・検証する。
'       不整合な日付（例：2月30日）を自動判定し、型安全な変換を提供する。
' 更新日：2026/03/24（MyLauncher統合・日付解析強化版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pRawValue As String
Private pYear     As Integer
Private pMonth    As Integer
Private pDay      As Integer
Private pIsValid  As Boolean

'===========================================================
' プロパティ Let/Get : RawValue
' 引数 : v - 解析対象文字列 ("20260324", "202603" 等)
' 機能 : 値をセットし、即座に解析（ParseValue）を実行する。
'===========================================================
Public Property Let RawValue(ByVal v As String)
    pRawValue = Trim(v)
    ParseValue
End Property

Public Property Get RawValue() As String: RawValue = pRawValue: End Property

' --- 解析結果プロパティ (ReadOnly) ---
Public Property Get IsValid() As Boolean: IsValid = pIsValid: End Property
Public Property Get Year() As Integer: Year = pYear: End Property
Public Property Get Month() As Integer: Month = pMonth: End Property
Public Property Get Day() As Integer: Day = pDay: End Property

'===========================================================
' プロパティ Get : Value
' 戻り値 : Date (不正な場合は 1900/01/01)
'===========================================================
Public Property Get Value() As Date
    If pIsValid Then
        Value = DateSerial(pYear, pMonth, pDay)
    Else
        Value = #1/1/1900#
    End If
End Property

'===========================================================
' 内部サブルーチン : ParseValue
' 機能 : 文字列長に応じて年・月・日を抽出。
'===========================================================
Private Sub ParseValue()
    pIsValid = False
    pYear = 0: pMonth = 0: pDay = 0
    
    ' 数字以外が含まれる場合は除外
    Dim cleanStr As String
    Dim i As Long, c As String
    For i = 1 To Len(pRawValue)
        c = Mid(pRawValue, i, 1)
        If c Like "[0-9]" Then cleanStr = cleanStr & c
    Next i
    
    On Error GoTo Err_Handler
    Select Case Len(cleanStr)
        Case 6 ' YYYYMM 形式
            pYear = CInt(Left(cleanStr, 4))
            pMonth = CInt(Right(cleanStr, 2))
            pDay = 1
            pIsValid = ValidateDate(pYear, pMonth, pDay)
            
        Case 8 ' YYYYMMDD 形式
            pYear = CInt(Left(cleanStr, 4))
            pMonth = CInt(Mid(cleanStr, 5, 2))
            pDay = CInt(Right(cleanStr, 2))
            pIsValid = ValidateDate(pYear, pMonth, pDay)
            
        Case Else
            pIsValid = False
    End Select
    Exit Sub

Err_Handler:
    pIsValid = False
End Sub

'===========================================================
' 内部関数 : ValidateDate
' 機能 : DateSerialの特性を利用して、実在する日付かチェックする。
'===========================================================
Private Function ValidateDate(y As Integer, m As Integer, d As Integer) As Boolean
    If y < 1900 Or y > 2100 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    
    Dim dt As Date
    dt = DateSerial(y, m, d)
    
    ' DateSerialは「2月30日」を入れると自動で「3月2日」等に変換してしまうため、
    ' 逆算して元に戻るかを確認することで、日付の正当性を厳密に判定する。
    ValidateDate = (Year(dt) = y And Month(dt) = m And Day(dt) = d)
End Function
