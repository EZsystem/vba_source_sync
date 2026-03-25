Attribute VB_Name = "mod_SkipCheck"
'~~~~~~~~~~~~~~ mod_SkipCheck ~~~~~~~~~~~~~~
Option Compare Database
Option Explicit
'-------------------------------------------------------
' ■ tdll_削除レコードからスキップ条件を読み込み
'    clsDeleteConditionクラスのインスタンスをCollectionに格納して返す
'-------------------------------------------------------
Public Function LoadSkipConditions() As Collection
    On Error GoTo ErrHandle
    
    Dim db As DAO.Database
    Dim rsDel As DAO.Recordset
    
    Dim colConditions As New Collection
    Dim cond As clsDeleteCondition
    
    Set db = CurrentDb()
    
    ' ※必要に応じて「演算子」列をSELECTするなど拡張可能
    '   今回は演算子列を持たない前提で、常に "Like" として格納
    Set rsDel = db.OpenRecordset("SELECT 対象フィールド名, 削除対象値 FROM tdll_削除レコード", dbOpenDynaset)
    
    Do While Not rsDel.EOF
        Set cond = New clsDeleteCondition
        cond.targetField = rsDel!対象フィールド名 & ""
        cond.TargetValue = rsDel!削除対象値 & ""
        cond.OperatorStr = "Like"  ' 基本運用; 必要に応じてテーブルに演算子列を追加
        
        colConditions.Add cond
        rsDel.MoveNext
    Loop
    
    rsDel.Close
    Set rsDel = Nothing
    db.Close
    Set db = Nothing
    
    Set LoadSkipConditions = colConditions
    Exit Function

ErrHandle:
    Debug.Print "LoadSkipConditionsでエラー発生: " & Err.description
    MsgBox "LoadSkipConditionsでエラー発生: " & Err.description, vbExclamation
    Set LoadSkipConditions = Nothing
End Function

'-------------------------------------------------------
' ■ 1レコードぶん(行ごと)のデータをチェックし、
'    スキップ対象なら True を返す。
'
'   [引数]
'       colConditions:  LoadSkipConditions()で取得したCollection
'       fieldNames():   タイトル行(B列～)の配列
'       rowValues():    実際のセルの値(B列～)の配列
'-------------------------------------------------------
Public Function ShouldSkipRow( _
    ByVal colConditions As Collection, _
    ByRef fieldNames() As Variant, _
    ByRef rowValues() As Variant) As Boolean
    
    On Error GoTo ErrHandle
    
    Dim cond As clsDeleteCondition
    Dim i As Long
    
    ShouldSkipRow = False
    
    ' スキップ条件をすべてチェック
    For Each cond In colConditions
        ' フィールド名に該当する列があるか探す
        For i = LBound(fieldNames) To UBound(fieldNames)
            If fieldNames(i) = cond.targetField Then
                
                Dim cellVal As String
                cellVal = Nz(rowValues(i), "")  ' Null対策
                
                Select Case cond.OperatorStr
                    Case "Like"
                        If cellVal Like cond.TargetValue Then
                            ShouldSkipRow = True
                            Exit For
                        End If
                    Case "="
                        If cellVal = cond.TargetValue Then
                            ShouldSkipRow = True
                            Exit For
                        End If
                    Case Else
                        ' デフォルトはLike扱いなど
                        If cellVal Like cond.TargetValue Then
                            ShouldSkipRow = True
                            Exit For
                        End If
                End Select
                
            End If
            
            If ShouldSkipRow Then Exit For
        Next i
        
        If ShouldSkipRow Then Exit For
    Next cond
    
    Exit Function
    
ErrHandle:
    Debug.Print "ShouldSkipRowでエラー発生: " & Err.description
    MsgBox "ShouldSkipRowでエラー発生: " & Err.description, vbExclamation
    ' エラー時はスキップ扱いにする等、要件に合わせて
    ShouldSkipRow = True
End Function


