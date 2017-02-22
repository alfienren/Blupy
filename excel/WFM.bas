Attribute VB_Name = "WFM"
Option Private Module
Option Explicit

Sub Process_Raw_Reports()

Dim rFloodlightCell         As Range
Dim rSAData                 As Range

Dim wSATemp                 As String

With Application
    
    .ScreenUpdating = True
    .EnableEvents = False
    .Calculation = xlCalculationManual
    
End With

wSATemp = "SA_Temp"

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

Sheets("Lookup").Activate

Range("AG1").Value = ThisWorkbook.FullName

End Sub

Sub Postprocess_Report()

Dim wSATemp                 As String

wSATemp = "SA_Temp"

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wSATemp).Delete
Err.Clear

Application.DisplayAlerts = True
Sheets("Lookup").Activate
Range("AG1").Clear
Sheets("Pivot").Activate

End Sub

Sub Call_WFM_Reporting()

RunPython ("import main; main.generate_wfm_reporting()")

End Sub

Sub WFM_Reporting()

Call Process_Raw_Reports
Call Postprocess_Report
Call Call_WFM_Reporting

End Sub