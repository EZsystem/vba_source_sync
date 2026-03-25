Attribute VB_Name = "mod_icubeValRunner"
'=================================================
' サブルーチン名 : Run_All_iCubeValidation_WithLog
' 説明   : iCubeバリデーション処理をログフォームに出力しながら順次実行する
'=================================================
Public Sub Run_All_iCubeValidation_WithLog()
    
    Call ClearLog
    Call AppendLog("iCubeバリデーション開始")

    ' ----------------------------
    ' フェーズ1：全体チェック
    ' ----------------------------
    Call AppendLog("処理1：mod_icube_All_1 開始")
    Call mod_icube_All_1
    Call AppendLog("処理1：mod_icube_All_1 終了")

    ' ----------------------------
    ' フェーズ2：Val2チェック全般
    ' ----------------------------
    Call AppendLog("処理2：mod_icube_Val2ALL 開始")
    Call mod_icube_Val2ALL
    Call AppendLog("処理2：mod_icube_Val2ALL 終了")

    ' ----------------------------
    ' フェーズ3：用途区分の補正
    ' ----------------------------
    Call AppendLog("処理3：Correct_CategoryUsage 開始")
    Call Correct_CategoryUsage
    Call AppendLog("処理3：Correct_CategoryUsage 終了")

    ' ----------------------------
    ' フェーズ4：価格カテゴリの割当
    ' ----------------------------
    Call AppendLog("処理4：assign_priceCategory 開始")
    Call assign_priceCategory
    Call AppendLog("処理4：assign_priceCategory 終了")

    ' ----------------------------
    ' フェーズ5：s基本工事コードとs基本工事名への転写1
    ' ----------------------------
    Call AppendLog("処理5：Run_FieldTranscribe_Default 開始")
    Call Run_FieldTranscribe_Default
    Call AppendLog("処理5：Run_FieldTranscribe_Default 終了")

    ' ----------------------------
    ' フェーズ6：s基本工事コードとs基本工事名への転写2
    ' ----------------------------
    Call AppendLog("処理6：Run_FieldTranscribe_WithSkipList 開始")
    Call Run_FieldTranscribe_WithSkipList
    Call AppendLog("処理6：Run_FieldTranscribe_WithSkipList 終了")

    ' ----------------------------
    ' フェーズ7：追加工事名称_cle処理
    ' ----------------------------
    Call AppendLog("処理7：Update_追加工事名称_Cle 開始")
    Call Update_追加工事名称_Cle
    Call AppendLog("処理7：Update_追加工事名称_Cle 終了")

    ' ----------------------------
    ' フェーズ8：顧客名の転写処理
    ' ----------------------------
    Call AppendLog("処理8：Transfer_顧客名_IfNotExists 開始")
    Call Transfer_顧客名_IfNotExists
    Call AppendLog("処理8：Transfer_顧客名_IfNotExists 終了")

    Call AppendLog("iCubeバリデーション完了")

End Sub

