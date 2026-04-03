VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xl_clsNameManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'簡易テスト済み
'=====================================
' Class Module: xl_clsNameManager
' 説明　：名前定義の追加・削除・一覧を管理するユーティリティクラスにゃ
' 作成日：2025/05/15
' 修正日：2025/05/15（Commit安全化・NameExists補強）
'=====================================
Option Explicit

' --- 内部保持変数 ---
Private pWorkbook As Workbook
Private pNameBuffer As Object ' Dictionary: name → range
Private pDeleteBuffer As Object ' Dictionary: name → dummy
Private pOverwrite As Boolean

'=================================================
' サブルーチン名 : Init
' 引数   : wb（省略可）- 操作対象ブック。省略時は ThisWorkbook
' 機能   : 初期化にゃ
'=================================================
Public Sub Init(Optional wb As Workbook)
    Set pWorkbook = IIf(wb Is Nothing, ThisWorkbook, wb)
    Set pNameBuffer = CreateObject("Scripting.Dictionary")
    Set pDeleteBuffer = CreateObject("Scripting.Dictionary")
    pOverwrite = True
End Sub

'=================================================
' プロパティ名 : OverwriteExisting
' 機能   : 上書きするかどうかのフラグにゃ
'=================================================
Public Property Let OverwriteExisting(val As Boolean)
    pOverwrite = val
End Property

Public Property Get OverwriteExisting() As Boolean
    OverwriteExisting = pOverwrite
End Property

'=================================================
' サブルーチン名 : AddName
' 引数   : nameStr, refRange
' 機能   : 登録キューに追加するにゃ
'=================================================
Public Sub AddName(nameStr As String, refRange As Range)
    If Trim(nameStr) = "" Then Exit Sub
    pNameBuffer(CStr(nameStr)) = refRange
End Sub

'=================================================
' サブルーチン名 : RemoveName
' 引数   : nameStr
' 機能   : 削除キューに登録するにゃ
'=================================================
Public Sub RemoveName(nameStr As String)
    If Trim(nameStr) = "" Then Exit Sub
    pDeleteBuffer(CStr(nameStr)) = True
End Sub

'=================================================
' サブルーチン名 : Commit
' 機能   : 登録・削除を実行するにゃ
'=================================================
Public Sub Commit()
    Dim key As Variant

    ' 削除処理
    For Each key In pDeleteBuffer.Keys
        If VarType(key) = vbString And Trim(key) <> "" Then
            On Error Resume Next
            pWorkbook.Names(CStr(key)).Delete
            On Error GoTo 0
        End If
    Next

    ' 追加処理
    For Each key In pNameBuffer.Keys
        If Trim(key) = "" Then GoTo SkipAdd

        If NameExists(CStr(key)) Then
            If pOverwrite Then
                On Error Resume Next
                pWorkbook.Names(CStr(key)).Delete
                On Error GoTo 0
            Else
                GoTo SkipAdd
            End If
        End If

        pWorkbook.Names.Add Name:=CStr(key), RefersTo:=pNameBuffer(key)

SkipAdd:
    Next
End Sub

'=================================================
' 関数名 : NameExists
' 説明   : 名前が既に存在するか確認するにゃ（スコープ付きにも対応）
'=================================================
Private Function NameExists(nameStr As String) As Boolean
    If Len(Trim(nameStr)) = 0 Then Exit Function

    Dim nm As Name
    For Each nm In pWorkbook.Names
        If Right(nm.Name, Len(nameStr)) = nameStr Then
            NameExists = True
            Exit Function
        End If
    Next
    NameExists = False
End Function

'=================================================
' 関数名 : ListDefinedNames
' 戻り値 : Dictionary（name → refersToFormula）
' 機能   : 現在定義されている名前一覧を返すにゃ
'=================================================
Public Function ListDefinedNames() As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim nm As Name
    For Each nm In pWorkbook.Names
        result.Add nm.Name, nm.RefersTo
    Next
    Set ListDefinedNames = result
End Function

'=================================================
' サブルーチン名 : Clean_XlfnNames
' 機能   : "_xlfn." を含む名前定義をすべて削除するにゃ
'=================================================
Public Sub Clean_XlfnNames()
    Dim nm As Name
    Dim deleteList As Collection: Set deleteList = New Collection

    ' 削除対象を収集（中で削除すると列挙バグるから分離にゃ）
    For Each nm In pWorkbook.Names
        If InStr(1, nm.Name, "_xlfn.", vbTextCompare) > 0 Then
            deleteList.Add nm.Name
        End If
    Next

    ' 削除実行
    Dim key As Variant
    For Each key In deleteList
        On Error Resume Next
        pWorkbook.Names(CStr(key)).Delete
        On Error GoTo 0
    Next
End Sub

