Attribute VB_Name = "xl_mod_IdMapping_Main"

'-------------------------------------
' Module: xl_mod_IdMapping_Main
' 説明　：高速化クラス(cls_IdMappingProcessor)の起動モジュール
'-------------------------------------
Option Explicit

Public Sub ExecuteIdMappingAndIdpaProcess()
    ' 完全再現版クラスを生成して実行
    Dim processor As New cls_IdMappingProcessor
    processor.ExecuteFullProcess
    
    MsgBox "全ての処理が完了しました。" & vbCrLf & _
           "（オリジナルロジック完全再現版）", vbInformation
End Sub

