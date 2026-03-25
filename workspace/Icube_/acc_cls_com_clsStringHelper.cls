VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "com_clsStringHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: com_clsStringHelper
' 説明：文字列の正規化、分割、結合、検索、置換を一括管理する。
'       NormalizeText により、全半角・空白の表記ゆれを吸収した比較が可能。
' 更新日：2026/03/24（MyLauncher統合・StringProcessor機能集約版）
'===========================================================
Option Explicit

'===========================================================
' 関数名 : NormalizeText
' 引数 : val - 対象文字列
' 戻り値 : 正規化された文字列
' 機能 : 1.前後の空白除去 2.全角英数・スペースを半角に統一 3.大文字統一
'===========================================================
Public Function NormalizeText(ByVal val As Variant) As String
    Dim s As String: s = Trim(CStr(Nz(val, "")))
    If s = "" Then
        NormalizeText = ""
        Exit Function
    End If
    
    ' 全角英数字、全角スペースを半角に変換 (vbNarrow)
    ' カタカナは全角のまま維持する場合が多いが、要件に応じ調整
    s = StrConv(s, vbNarrow)
    
    ' 大文字に統一して比較しやすくする
    NormalizeText = UCase(s)
End Function

'===========================================================
' 関数名 : ContainsText
' 引数 : source - 検索対象
'        keyword - キーワード
'        useNormalization - 正規化して比較するか (Default: True)
' 戻り値 : Boolean
' 機能 : 部分一致判定。正規化を併用することで表記ゆれに強い検索を行う。
'===========================================================
Public Function ContainsText(ByVal source As String, _
                             ByVal keyword As String, _
                             Optional ByVal useNormalization As Boolean = True) As Boolean
    
    Dim s As String, k As String
    If useNormalization Then
        s = Me.NormalizeText(source)
        k = Me.NormalizeText(keyword)
    Else
        s = source
        k = keyword
    End If
    
    If k = "" Then
        ContainsText = False
    Else
        ContainsText = (InStr(1, s, k, vbTextCompare) > 0)
    End If
End Function

'===========================================================
' 関数名 : SplitToArray / JoinArray
' 機能 : 文字列と配列の相互変換（標準機能のラップ）
'===========================================================
Public Function SplitToArray(ByVal s As String, ByVal delimiter As String) As Variant
    SplitToArray = Split(s, delimiter)
End Function

Public Function JoinArray(ByRef arr As Variant, ByVal delimiter As String) As String
    JoinArray = Join(arr, delimiter)
End Function

'===========================================================
' 関数名 : ReplaceAll
' 引数 : source - 元の文字列
'        find - 検索文字列
'        replace - 置換文字列
'===========================================================
Public Function ReplaceAll(ByVal source As String, ByVal find As String, ByVal replace As String) As String
    ReplaceAll = replace(source, find, replace)
End Function

'===========================================================
' 関数名 : ExtractNumeric
' 引数 : val - 文字列
' 戻り値 : 数字のみを抽出した文字列
'===========================================================
Public Function ExtractNumeric(ByVal val As String) As String
    Dim i As Long, char As String, result As String
    For i = 1 To Len(val)
        char = Mid(val, i, 1)
        If char Like "[0-9]" Then result = result & char
    Next i
    ExtractNumeric = result
End Function

' --- 旧 StringProcessor 互換用メソッド ---
Public Function ConvertToUpper(ByVal s As String) As String
    ConvertToUpper = UCase(s)
End Function
