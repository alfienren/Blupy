Attribute VB_Name = "Process_Raw"
Option Explicit

Sub ProcessRawDFA()

Dim rFloodlightCell         As Range
Dim rSAData                 As Range
Dim rCFVData                As Range

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
    .Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
End With

Sheets("CFV").Activate

With ActiveSheet

    Set rFloodlightCell = Cells.Find(What:="Floodlight Attribution Type", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    
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
    
    Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
                
End With

End Sub
