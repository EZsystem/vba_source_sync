VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Launcher 
   Caption         =   "業務実行パネル"
   ClientHeight    =   6780
   ClientLeft      =   -36
   ClientTop       =   -120
   ClientWidth     =   8400
   OleObjectBlob   =   "xl_frm_Launcher.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frm_Launcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==========================================
' モジュール：frm_Launcher
' 説明：3区分(アルファベット/システム/他)対応型ランチャー
' ==========================================

Private pMenuArray As Variant ' メニュー配列(0:表示名, 1:ジャンプ先/マクロ, 2:説明, 3:前処理)

' ------------------------------------------
'  1. 初期化・終了処理
' ------------------------------------------

Private Sub UserForm_Initialize()
    Me.MultiPage1.value = 0
    
    ' 区分コンボボックスの初期化
    With Me.cmb_Group
        .Clear
        .AddItem "すべて"
        .AddItem "アルファベット"
        .AddItem "システム"
        .AddItem "その他"
        .ListIndex = 0
    End With
    
    ' 初期状態では非表示
    Me.cmb_Group.Visible = False
    Me.cmb_AlphaFilter.Visible = False
    
    Call LoadMenu_Integrated(0)
End Sub

Private Sub btn_Close_Click()
    Unload Me
End Sub

' ------------------------------------------
'  2. イベントハンドラ
' ------------------------------------------

Private Sub MultiPage1_Change()
    Dim isJumpTab As Boolean: isJumpTab = (Me.MultiPage1.value = 5)
    
    ' ジャンプタブ(5)の時だけフィルタ系を表示
    Me.cmb_Group.Visible = isJumpTab
    
    ' アルファベット選択時以外は頭文字フィルタを隠す
    If isJumpTab Then
        Call RefreshAlphaVisibility
        Call LoadExistingInitials
    Else
        Me.cmb_AlphaFilter.Visible = False
    End If
    
    Call LoadMenu_Integrated(Me.MultiPage1.value)
End Sub

''' <summary>
''' 区分（グループ）変更時
''' </summary>
Private Sub cmb_Group_Change()
    If Me.MultiPage1.value = 5 Then
        Call RefreshAlphaVisibility
        Call LoadMenu_Integrated(5)
    End If
End Sub

''' <summary>
''' 頭文字フィルタ変更時
''' </summary>
Private Sub cmb_AlphaFilter_Change()
    If Me.MultiPage1.value = 5 Then Call LoadMenu_Integrated(5)
End Sub

Private Sub lst_Menu_Click()
    Dim idx As Integer: idx = Me.lst_Menu.ListIndex
    If idx < 0 Or IsEmpty(pMenuArray) Then Exit Sub
    Me.txt_Description.Text = "【概要】" & vbCrLf & pMenuArray(idx)(2) & _
                              IIf(pMenuArray(idx)(3) <> "", vbCrLf & vbCrLf & "【前処理】" & vbCrLf & pMenuArray(idx)(3), "")
End Sub

Private Sub btn_Execute_Click()
    Dim idx As Integer: idx = Me.lst_Menu.ListIndex
    If idx < 0 Then Exit Sub
    
    ' 前処理確認
    If pMenuArray(idx)(3) <> "" Then
        If MsgBox(pMenuArray(idx)(3), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    If Me.MultiPage1.value = 5 Then
        On Error Resume Next
        ThisWorkbook.Worksheets(pMenuArray(idx)(1)).Activate
        On Error GoTo 0
    Else
        Call xl_mod_Launcher_Local.Launcher_Execute_Macro(pMenuArray(idx)(1))
    End If
End Sub

' ------------------------------------------
'  3. 内部ロジック
' ------------------------------------------

''' <summary>
''' アルファベットグループが選ばれている時だけ、頭文字フィルタを出す
''' </summary>
Private Sub RefreshAlphaVisibility()
    Me.cmb_AlphaFilter.Visible = (Me.cmb_Group.value = "アルファベット")
End Sub

''' <summary>
''' 大文字・小文字を区別して存在する頭文字を抽出
''' </summary>
Private Sub LoadExistingInitials()
    Dim ws As Worksheet, dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Me.cmb_AlphaFilter.Clear
    Me.cmb_AlphaFilter.AddItem "すべて"
    
    For Each ws In ThisWorkbook.Worksheets
        Dim s As String: s = Left(ws.Name, 1)
        ' 純粋なアルファベット（A-Z, a-z）のみ抽出
        If (Asc(s) >= 65 And Asc(s) <= 90) Or (Asc(s) >= 97 And Asc(s) <= 122) Then
            dict(s) = True
        End If
    Next ws
    
    ' 辞書にある文字を追加
    Dim key As Variant
    For Each key In dict.keys
        Me.cmb_AlphaFilter.AddItem key
    Next key
    Me.cmb_AlphaFilter.ListIndex = 0
End Sub

''' <summary>
''' 統合ロードロジック
''' </summary>
Private Sub LoadMenu_Integrated(ByVal pageIndex As Integer)
    Me.lst_Menu.Clear
    Dim tempList() As Variant: Dim count As Long: count = 0
    
    ' --- Case 5: 3区分対応型シートスキャン ---
    If pageIndex = 5 Then
        Dim ws As Worksheet
        Dim gFilter As String: gFilter = Me.cmb_Group.value
        Dim aFilter As String: aFilter = Me.cmb_AlphaFilter.value
        
        For Each ws In ThisWorkbook.Worksheets
            Dim first As String: first = Left(ws.Name, 1)
            Dim category As String
            
            ' カテゴリ判定
            If first = "_" Then
                category = "システム"
            ElseIf (Asc(first) >= 65 And Asc(first) <= 90) Or (Asc(first) >= 97 And Asc(first) <= 122) Then
                category = "アルファベット"
            Else
                category = "その他"
            End If
            
            ' フィルタ照合
            Dim matchGroup As Boolean: matchGroup = (gFilter = "すべて" Or gFilter = category)
            Dim matchAlpha As Boolean: matchAlpha = True
            If category = "アルファベット" And aFilter <> "すべて" And aFilter <> "" Then
                matchAlpha = (first = aFilter) ' 大文字小文字を厳密に判定
            End If
            
            If matchGroup And matchAlpha Then
                ReDim Preserve tempList(count)
                tempList(count) = Array(ws.Name, ws.Name, "シート '" & ws.Name & "' にジャンプします。", "")
                count = count + 1
            End If
        Next ws
        GoTo Finalize
    End If

    ' --- Case 0-4: テーブル読込 (ID順=テーブル並び順) ---
    Dim lo As ListObject: Set lo = ThisWorkbook.Worksheets("_LauncherMenu").ListObjects("xt_LauncherMenu")
    Dim data As Variant: data = lo.DataBodyRange.value
    Dim tabCap As String: tabCap = Me.MultiPage1.Pages(pageIndex).Caption
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If data(i, 2) = tabCap And data(i, 8) = 1 Then
            ReDim Preserve tempList(count)
            tempList(count) = Array(data(i, 3), data(i, 5) & "." & data(i, 4), data(i, 6), data(i, 7))
            count = count + 1
        End If
    Next i

Finalize:
    pMenuArray = tempList
    If count > 0 Then
        For i = 0 To count - 1
            Me.lst_Menu.AddItem pMenuArray(i)(0)
        Next i
    End If
End Sub

