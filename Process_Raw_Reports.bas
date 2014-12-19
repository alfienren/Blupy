Attribute VB_Name = "Process_Raw_Reports"
Option Explicit

Sub ProcessRawDFA()

Dim First                   As Long
Dim Last                    As Long
Dim Lrow                    As Long

Dim rUniqueID               As Range
Dim rFloodlightCell         As Range

Dim rSAData                 As Range
Dim rUniqueIDSA             As Range

Dim rUniqueIDCFV            As Range
Dim rCFVData                As Range
Dim rCFVCell                As Range

Dim cell                    As Range
Dim cell2                   As Range
Dim var                     As Range

Dim wCFVTemp                As String
Dim wSATemp                 As String

With Application
    
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
    
End With

wCFVTemp = "CFV_Temp"
wSATemp = "SA_Temp"

Sheets("SA").Activate

With ActiveSheet

    Set rUniqueID = Cells.Find(What:="Unique ID", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        
    If rUniqueID Is Nothing Then
    
    Else
    
        rUniqueID.Columns.Delete Shift:=xlToLeft
        
    End If
    
    .Range("C1").Select
    Selection.End(xlDown).Select
    Range(Selection.End(xlToRight), Selection.End(xlToLeft)).Select
    
    Set rSAData = Range(Selection, Selection.End(xlDown))

End With

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wSATemp).Delete
Err.Clear

Application.DisplayAlerts = True
Worksheets.Add.Name = wSATemp

Sheets("SA_Temp").Activate

With ActiveSheet

    Cells.ClearContents
    rSAData.Copy
    .Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    .Range("B1").End(xlDown).EntireRow.Delete
    
    .Range("A1").Value = "UniqueID"
    .Range("A2").Select
    
    Range(Selection, Selection.Offset(rSAData.Rows.Count - 3, 0)).Select
    
    Selection.Formula = "=RC[1]&RC[2]&RC[3]&RC[9]&RC[12]"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    
    
End With

Sheets("CFV").Activate

With ActiveSheet

    Set rFloodlightCell = Cells.Find(What:="Floodlight Attribution Type", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)

    Set rUniqueID = Cells.Find(What:="Unique ID", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        
    If rUniqueID Is Nothing Then
    
    Else
    
        rUniqueID.EntireColumn.Delete Shift:=xlToLeft
        
    End If
    
    rFloodlightCell.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    Set rCFVData = Range(Selection, Selection.End(xlDown).Offset(-1, 0))
    
End With

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wCFVTemp).Delete
Err.Clear

Application.DisplayAlerts = True
Worksheets.Add.Name = wCFVTemp
    
rCFVData.Copy

Sheets("CFV_Temp").Activate
With ActiveSheet
    
    Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("A1").Value = "UniqueID"

    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Offset(0, -1).Select

    Selection.Formula = "=RC[1]&RC[2]&RC[3]&RC[9]&RC[11]"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    Range("A1").CurrentRegion.Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
            
End With

End Sub
