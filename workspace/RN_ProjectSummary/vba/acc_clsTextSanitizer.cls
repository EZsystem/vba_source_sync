VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTextSanitizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
' Class Module: acc_clsTextSanitizer
' 説明：文字列データの安全性確保、制御文字・不可視文字の除去を行う。
'       Accessのインポート一時テーブルや配列データの正規化に特化。
' 更新日：2026/03/24（MyLauncher統合・高速洗浄版）
'===========================================================
Option Explicit

' --- 内部保持変数 ---
Private pErr As com_clsErrorUtility

'===========================================================
' サブルーチン名 : Init
' 引数 : errorUtil - (任意) com_clsErrorUtility
' 機能 : 依存オブジェクトを初期化する。
'===========================================================
Public Sub Init(Optional ByRef errorUtil As com_clsErrorUtility)
    If Not errorUtil Is Nothing Then
        Set pErr = errorUtil
    End If
End Sub

'===========================================================
' 関数名 : CleanText
' 引数 : value - 変換対象
'        defaultValue - Null時の戻り値
' 戻り値 : String (Trim済み、改行・制御文字除去済み)
'===========================================================
Public Function CleanText(ByVal Value As Variant, _
                          Optional ByVal defaultValue As String = "") As String
    If IsNull(Value) Or IsEmpty(Value) Then
        CleanText = defaultValue
        Exit Function
    End If
    
    Dim s As String: s = CStr(Value)
    
    ' 制御文字・改行の一括除去
    s = Replace(s, vbCrLf, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, Chr(160), "") ' ノンブレークスペース対策
    
    CleanText = Trim(s)
End Function

'===========================================================
' 関数名 : RemoveAllSpaces
' 機能 : 文字列内の全スペース（全角・半角）を除去する。
'===========================================================
Public Function RemoveAllSpaces(ByVal Value As String) As String
    Dim s As String
    s = Replace(Value, " ", "")  ' 半角
    s = Replace(s, "　", "") ' 全角
    RemoveAllSpaces = s
End Function

'===========================================================
' 関数名 : IsEmptyText
' 機能 : Null、空文字、空白のみの文字列を判定する。
'===========================================================
Public Function IsEmptyText(ByVal Value As Variant) As Boolean
    If IsNull(Value) Then
        IsEmptyText = True
    Else
        IsEmptyText = (Trim(CStr(Value)) = "")
    End If
End Function
