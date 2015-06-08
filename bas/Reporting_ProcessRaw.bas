Attribute VB_Name = "Reporting_Process_Raw"
Option Explicit

Sub Process_Raw_Reports()

Dim rFloodlightCell         As Range
Dim rSAData                 As Range
Dim rCFVData                As Range

Dim wCFVTemp                As String
Dim wSATemp                 As String
Dim wDDR                    As String
Dim wSummary                As String

With Application
    
    .ScreenUpdating = True
    .EnableEvents = False
    .Calculation = xlCalculationManual
    
End With

wCFVTemp = "CFV_Temp"
wSATemp = "SA_Temp"
wDDR = "DDR"
wSummary = "Summary"

Sheets("SA").Activate

With ActiveSheet
    
    .Range("C1").Select
    Selection.End(xlDown).Select
    Range(Selection.End(xlToRight), Selection.End(xlToLeft)).Select
    
    Set rSAData = Range(Selection, Selection.End(xlDown).Offset(-1, 0))

End With

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wSATemp).Delete
Worksheets(wDDR).Delete
Worksheets(wSummary).Delete
Err.Clear

Application.DisplayAlerts = True
Worksheets.Add.Name = wSATemp

Sheets("SA_Temp").Activate

With ActiveSheet

    Cells.ClearContents
    rSAData.Copy
    .Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
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
    
    Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
                
End With

End Sub
