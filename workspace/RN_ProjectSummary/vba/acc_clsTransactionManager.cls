VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsTransactionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsTransactionManager
' 隱ｬ譏弱��・咼AO繝医Λ繝ｳ繧ｶ繧ｯ繧ｷ繝ｧ繝ｳ縺ｮ髢句ｧ九・繧ｳ繝溘ャ繝医・繝ｭ繝ｼ繝ｫ繝舌ャ繧ｯ繧堤ｮ｡逅・☆繧九け繝ｩ繧ｹ
'=====================================
Option Explicit

' --- 蜀・Κ菫晄戟螟画焚 ---
Private pInTransaction As Boolean

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: Init
'=================================================
Public Sub Init(Optional db As DAO.Database = Nothing)
    pInTransaction = False
End Sub

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: BeginTransaction
' 讖溯・ : 繝医Λ繝ｳ繧ｶ繧ｯ繧ｷ繝ｧ繝ｳ繧帝幕蟋九☆繧九�・AO縺ｮ豁｣縺励＞髫主ｱ､・・orkspace・峨ｒ蜻ｼ縺ｳ蜃ｺ縺励∪縺・'=================================================
Public Sub BeginTransaction()
    If Not pInTransaction Then
        ' DAO縺ｮ莉墓ｧ假ｼ壹ヨ繝ｩ繝ｳ繧ｶ繧ｯ繧ｷ繝ｧ繝ｳ縺ｯ Workspace(0) 縺ｧ髢句ｧ九＠縺ｾ縺・        DBEngine.Workspaces(0).BeginTrans
        pInTransaction = True
    End If
End Sub

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: Commit
' 讖溯・ : 繝医Λ繝ｳ繧ｶ繧ｯ繧ｷ繝ｧ繝ｳ繧偵さ繝溘ャ繝医＠縺ｦ遒ｺ螳壹☆繧・'=================================================
Public Sub Commit()
    If pInTransaction Then
        DBEngine.Workspaces(0).CommitTrans
        pInTransaction = False
    End If
End Sub

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: Rollback
' 讖溯・ : 繝医Λ繝ｳ繧ｶ繧ｯ繧ｷ繝ｧ繝ｳ繧偵Ο繝ｼ繝ｫ繝舌ャ繧ｯ縺励※蜈・↓謌ｻ縺・'=================================================
Public Sub Rollback()
    If pInTransaction Then
        ' DAO縺ｮ繝｡繧ｽ繝・ラ蜷阪・ Rollback 縺ｧ縺呻ｼ・ollbackTrans縺ｧ縺ｯ縺ゅｊ縺ｾ縺帙ｓ・・        DBEngine.Workspaces(0).Rollback
        pInTransaction = False
    End If
End Sub
