VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acc_clsExcelImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
' Class Module: acc_clsExcelImporter (EZsystem Refactored)
' 隱ｬ譏弱��・哘xcel驟榊・縺九ｉAccess繝・・繝悶Ν縺ｸ縺ｮ鬮倬�溷叙霎ｼ繧呈署萓帙☆繧九け繝ｩ繧ｹ
' 譖ｴ譁ｰ譌･・・026/03/20 - 繝励Ο繝代ユ繧｣蜷阪・邨ｱ荳�縲∝・蜷肴ｭ｣隕丞喧縺ｮ螳溯｣・'=====================================
Option Explicit

' --- 蜀・Κ菫晄戟螟画焚 ---
Private pDB               As DAO.Database
Private pDictMapping      As Object ' 繧ｿ繧､繝医Ν -> 繝輔ぅ繝ｼ繝ｫ繝牙錐
Private pColIndex         As Object ' 繧ｿ繧､繝医Ν -> 驟榊・蛻礼分蜿ｷ
Private pSkipConditions   As Object
Private pDefaultValues    As Object

' --- 蜈ｬ髢句､画焚・郁ｨｭ螳夂畑・・---
Public TargetTable     As String    ' 譛ｬ逡ｪ繝・・繝悶Ν蜷搾ｼ・ableName縺ｮ繧ｨ繧､繝ｪ繧｢繧ｹ・・Public TempTable       As String    ' 荳�譎ゅユ繝ｼ繝悶Ν蜷搾ｼ・empTableName縺ｮ繧ｨ繧､繝ｪ繧｢繧ｹ・・Public TitleRow        As Long      ' 繧ｿ繧､繝医Ν陦鯉ｼ亥・譛溷�､ 1・・Public DataStartRow    As Long      ' 繝・・繧ｿ髢句ｧ玖｡鯉ｼ亥・譛溷�､ 2・・Public StartColumn     As Long      ' 髢句ｧ句・・亥・譛溷�､ 1・・Public ImportColumnMax As Long      ' 譛�螟ｧ蛻玲焚・・縺ｮ蝣ｴ蜷医・閾ｪ蜍募愛螳夲ｼ・
' --- 莉･蜑阪・繧ｳ繝ｼ繝峨→縺ｮ莠呈鋤逕ｨ繝励Ο繝代ユ繧｣ ---
Public Property Let tableName(ByVal val As String): TargetTable = val: End Property
Public Property Let TempTableName(ByVal val As String): TempTable = val: End Property

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: Init
' 讖溯・ : 蜷・ｨｮ霎樊嶌繧ｪ繝悶ず繧ｧ繧ｯ繝医→DB謗･邯壹ｒ蛻晄悄蛹悶☆繧・'=================================================
Public Sub Init()
    Set pDB = CurrentDb
    Set pDictMapping = CreateObject("Scripting.Dictionary")
    Set pColIndex = CreateObject("Scripting.Dictionary")
    Set pSkipConditions = CreateObject("Scripting.Dictionary")
    Set pDefaultValues = CreateObject("Scripting.Dictionary")
    
    ' 繝・ヵ繧ｩ繝ｫ繝亥�､險ｭ螳・    TitleRow = 1
    DataStartRow = 2
    StartColumn = 1
End Sub

'=================================================
' 繝励Ο繝代ユ繧｣ : MappingMode
' 讖溯・ : "繧ｿ繧､繝医Ν荳�閾ｴ" 縺ｮ蝣ｴ蜷医�∬・蜍輔〒繝槭ャ繝斐Φ繧ｰ繧堤函謌舌☆繧・'=================================================
Public Property Let MappingMode(ByVal mode As String)
    If mode = "繧ｿ繧､繝医Ν荳�閾ｴ" Then
        Dim key As Variant
        For Each key In pColIndex.Keys
            pDictMapping(key) = key
        Next key
    End If
End Property

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: PrepareFieldNames
' 蠑墓焚 : titleRowArr - 繧ｿ繧､繝医Ν陦後・2谺｡蜈・・蛻・' 讖溯・ : 繧ｿ繧､繝医Ν繧呈ｭ｣隕丞喧縺励※蛻嶺ｽ咲ｽｮ繧定ｨ俶・縺吶ｋ・郁｡ｨ險倥ｆ繧悟ｯｾ蠢懶ｼ・'=================================================
Public Sub PrepareFieldNames(ByVal titleRowArr As Variant)
    Dim i As Long
    Dim rawTitle As String
    Dim normTitle As String
    
    pColIndex.RemoveAll
    
    ' 驟榊・縺ｮ蛻玲焚縺ｯ UBound(titleRowArr, 2)
    For i = 1 To UBound(titleRowArr, 2)
        rawTitle = Nz(titleRowArr(1, i), "")
        If rawTitle <> "" Then
            ' EZsystem讓呎ｺ厄ｼ壽ｭ｣隕丞喧縺励※逋ｻ骭ｲ
            normTitle = Normalize_Text(rawTitle)
            If Not pColIndex.Exists(normTitle) Then
                pColIndex.Add normTitle, i
            End If
        End If
    Next i
End Sub

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: ImportFromArray
' 蠑墓焚 : dataArr - 蜿冶ｾｼ蟇ｾ雎｡縺ｮ邏皮ｲ九↑繝・・繧ｿ驟榊・
' 讖溯・ : 鬮倬�溘↓ Recordset 縺ｸ豬√＠霎ｼ繧�
'=================================================
Public Sub ImportFromArray(ByVal dataArr As Variant)
    Dim rs As DAO.Recordset
    Dim r As Long
    Dim titleKey As Variant
    Dim targetField As String
    Dim valCell As Variant
    
    If TempTable = "" Then Err.Raise 999, , "TempTable縺梧欠螳壹＆繧後※縺・∪縺帙ｓ"
    
    Set rs = pDB.OpenRecordset(TempTable, dbOpenDynaset)
    
    On Error GoTo ErrLine
    For r = 1 To UBound(dataArr, 1)
        If Not ShouldSkip(dataArr, r) Then
            rs.AddNew
            ' 繝槭ャ繝斐Φ繧ｰ螳夂ｾｩ縺輔ｌ縺ｦ縺・ｋ蛻励・縺ｿ繝ｫ繝ｼ繝・            For Each titleKey In pDictMapping.Keys
                ' 繧ｿ繧､繝医Ν縺窪xcel蛻励↓蟄伜惠縺吶ｋ縺狗｢ｺ隱・                If pColIndex.Exists(titleKey) Then
                    targetField = pDictMapping(titleKey)
                    valCell = dataArr(r, pColIndex(titleKey))
                    
                    ' 遨ｺ逋ｽ譎ゅ・繝・ヵ繧ｩ繝ｫ繝亥�､陬懷ｮ・                    If (IsNull(valCell) Or valCell = "") Then
                        If pDefaultValues.Exists(targetField) Then
                            valCell = pDefaultValues(targetField)
                        End If
                    End If
                    
                    rs.Fields(targetField).Value = valCell
                End If
            Next titleKey
            rs.Update
        End If
    Next r
    
    rs.Close
    Exit Sub

ErrLine:
    MsgBox "陦・" & r & " 縺ｮ蜿冶ｾｼ荳ｭ縺ｫ繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: " & Err.Description, vbCritical
    If Not rs Is Nothing Then rs.Close
End Sub

'=================================================
' 髢｢謨ｰ蜷・: ShouldSkip (蜀・Κ逕ｨ)
'=================================================
Private Function ShouldSkip(dataArr As Variant, r As Long) As Boolean
    ' 蠢・ｦ√↓蠢懊§縺ｦ繧ｹ繧ｭ繝・・繝ｭ繧ｸ繝・け繧貞ｮ溯｣・ｼ育樟迥ｶ縺ｯ蜈ｨ莉ｶ騾壹☆・・    ShouldSkip = False
End Function

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: CommitToMain
' 讖溯・ : 荳�譎ゅユ繝ｼ繝悶Ν縺九ｉ譛ｬ逡ｪ繝・・繝悶Ν縺ｸ荳�諡ｬ霑ｽ蜉�
'=================================================
Public Sub CommitToMain()
    If TargetTable = "" Or TempTable = "" Then Exit Sub
    Dim sql As String
    sql = "INSERT INTO [" & TargetTable & "] SELECT * FROM [" & TempTable & "];"
    pDB.Execute sql, dbFailOnError
End Sub

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: ImportFromWorksheet
' 蠑墓焚 : ws - Excel縺ｮ繝ｯ繝ｼ繧ｯ繧ｷ繝ｼ繝医が繝悶ず繧ｧ繧ｯ繝・'=================================================
Public Sub ImportFromWorksheet(ByVal ws As Object)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim titleArr As Variant
    Dim dataArr As Variant
    
    ' 蛻玲焚縺ｮ閾ｪ蜍募愛螳・    If ImportColumnMax = 0 Then
        lastCol = ws.Cells(TitleRow, ws.Columns.count).End(-4159).Column ' xlToLeft
        ImportColumnMax = lastCol - StartColumn + 1
    End If
    
    ' 繧ｿ繧､繝医Ν蜿門ｾ励→隗｣譫・    titleArr = ws.Range(ws.Cells(TitleRow, StartColumn), ws.Cells(TitleRow, StartColumn + ImportColumnMax - 1)).value
    Call PrepareFieldNames(titleArr)
    
    ' 譛�邨り｡悟叙蠕・    lastRow = ws.Cells(ws.rows.Count, StartColumn).End(-4162).Row ' xlUp
    If lastRow < DataStartRow Then Exit Sub ' 繝・・繧ｿ縺ｪ縺・
    ' 繝・・繧ｿ荳�諡ｬ蜿門ｾ・    dataArr = ws.Range(ws.Cells(DataStartRow, StartColumn), ws.Cells(lastRow, StartColumn + ImportColumnMax - 1)).value
    
    ' 驟榊・縺九ｉ蜿冶ｾｼ
    Call ImportFromArray(dataArr)
End Sub

'=================================================
' 繧ｵ繝悶Ν繝ｼ繝√Φ蜷・: ImportUnpivotedData (New!)
' 蠑墓焚 : lo - ListObject
'        worksiteName - 菴懈･ｭ謇�蜷搾ｼ亥崋螳壼�､縺ｨ縺励※霑ｽ蜉�・・'        sourcePath - 蜈・ヵ繧｡繧､繝ｫ繝代せ・亥崋螳壼�､縺ｨ縺励※霑ｽ蜉�・・' 讖溯・ : 讓ｪ謖√■繝・・繧ｿ繧堤ｸｦ謖√■縺ｫ螟画鋤縺励↑縺後ｉ鬮倬�溘う繝ｳ繝昴・繝・'=================================================
Public Sub ImportUnpivotedData( _
    ByVal lo As Object, _
    ByVal worksiteName As String, _
    ByVal sourcePath As String)
    
    Dim rs      As DAO.Recordset
    Dim dataArr As Variant
    Dim r As Long, c As Long
    Dim colName As String
    Dim colNorm As String
    
    ' 蝗ｺ螳壼・縺ｮ繧､繝ｳ繝・ャ繧ｯ繧ｹ蜿門ｾ暦ｼ・ormalize_Text繧剃ｽｿ逕ｨ・・    Dim idxNo As Long:    idxNo = Get_ColumnIndex_Robust(lo, "No")
    Dim idxYear As Long:  idxYear = Get_ColumnIndex_Robust(lo, "蟷ｴ譛・)
    Dim idxCode As Long:  idxCode = Get_ColumnIndex_Robust(lo, "蟾･莠九さ繝ｼ繝・)
    Dim idxName As Long:  idxName = Get_ColumnIndex_Robust(lo, "蟾･莠句錐")
    Dim idxComm As Long:  idxComm = Get_ColumnIndex_Robust(lo, "繧ｳ繝｡繝ｳ繝・)
    
    If TempTable = "" Then Err.Raise 999, , "TempTable縺梧欠螳壹＆繧後※縺・∪縺帙ｓ"
    
    dataArr = lo.DataBodyRange.Value
    Set rs = pDB.OpenRecordset(TempTable, dbOpenDynaset)
    
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To lo.ListColumns.count
            colName = lo.ListColumns(c).Name
            colNorm = Normalize_Text(colName)
            
            ' 繝｡繧ｿ繝・・繧ｿ蛻励ｄ遨ｺ蜷阪・蛻励ｒ髯､螟・            If InStr(colName, vbLf) = 0 And _
               UCase(colNorm) <> "NO" And _
               colNorm <> Normalize_Text("蟷ｴ譛・) And _
               colNorm <> Normalize_Text("蟾･莠九さ繝ｼ繝・) And _
               colNorm <> Normalize_Text("蟾･莠句錐") And _
               colNorm <> Normalize_Text("繧ｳ繝｡繝ｳ繝・) And _
               colName <> "" Then
               
                ' 蛟､縺後≠繧句�ｴ蜷医・縺ｿ繝ｬ繧ｳ繝ｼ繝我ｽ懈・
                If Not IsEmpty(dataArr(r, c)) And dataArr(r, c) <> 0 Then
                    rs.AddNew
                    
                    ' 蝗ｺ螳壽ュ蝣ｱ
                    rs!蜈・ヵ繧｡繧､繝ｫ繝代せ = sourcePath
                    rs!菴懈･ｭ謇�蜷・= worksiteName
                    
                    ' 蜈・ョ繝ｼ繧ｿ縺ｮ蝗ｺ螳壹・繝・ヴ繝ｳ繧ｰ
                    If idxNo > 0 Then rs!No = dataArr(r, idxNo)
                    If idxYear > 0 Then rs!蟷ｴ譛・= dataArr(r, idxYear)
                    If idxCode > 0 Then rs!蟾･莠九さ繝ｼ繝・= dataArr(r, idxCode)
                    If idxName > 0 Then rs!蟾・莠句錐 = dataArr(r, idxName)
                    If idxComm > 0 Then rs!繧ｳ繝｡繝ｳ繝・= dataArr(r, idxComm)
                    
                    ' 螻暮幕諠・�ｱ・育､ｾ蜩｡蜷阪→蜑ｲ蜷茨ｼ・                    rs!遉ｾ蜩｡蜷・= colName
                    rs!蜈ｼ蜍咏紫蜑ｲ蜷・= dataArr(r, c)
                    
                    rs.Update
                End If
            End If
        Next c
    Next r
    
    rs.Close
End Sub
