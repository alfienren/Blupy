
Option Private Module
Option Explicit

Sub TMO_Raw_Reports()

Dim rFloodlightCell         As Range
Dim rSAData                 As Range
Dim rCFVData                As Range

Dim wCFVTemp                As String
Dim wSATemp                 As String
Dim wDDR                    As String
Dim wSummary                As String
Dim wQA                     As String

With Application

    .ScreenUpdating = True
    .EnableEvents = False
    .Calculation = xlCalculationManual

End With

wCFVTemp = "CFV_Temp"
wSATemp = "SA_Temp"
wDDR = "DDR"
wSummary = "Summary"
wQA = "Data_QA_Output"

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
Worksheets(wQA).Delete
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


Sub TMO_Postprocess_Report()

Dim wCFVTemp                As String
Dim wSATemp                 As String
Dim wSummary                As String
Dim wDDR                    As String

wCFVTemp = "CFV_Temp"
wSATemp = "SA_Temp"
wDDR = "DDR"
wSummary = "Summary"

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wSATemp).Delete
Worksheets(wCFVTemp).Delete
Worksheets(wDDR).Delete
Worksheets(wSummary).Delete

Err.Clear

Application.DisplayAlerts = True
Sheets("Lookup").Activate
Range("AA1").Clear
Sheets("Pivot").Activate

End Sub

Sub METRO_Raw_Reports()

Dim rFloodlightCell         As Range
Dim rSAData                 As Range
Dim rCFVData                As Range

Dim wCFVTemp                As String
Dim wSATemp                 As String
Dim wMissing                As String
Dim wQA                     As String

With Application

    .ScreenUpdating = True
    .EnableEvents = False
    .Calculation = xlCalculationManual

End With

wCFVTemp = "CFV_Temp"
wSATemp = "SA_Temp"
wMissing = "Placements w.o Categories"
wQA = "Data_QA_Output"

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
Worksheets(wQA).Delete
Err.Clear

Application.DisplayAlerts = True
Worksheets.Add.Name = wSATemp

Sheets("SA_Temp").Activate

With ActiveSheet

    Cells.ClearContents
    rSAData.Copy
    .Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, transpose:=False

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
        :=False, transpose:=False

End With

Sheets("Action_Reference").Activate
Range("AG1").Value = ActiveWorkbook.FullName

End Sub

Sub METRO_Postprocess_Report()

Dim wCFVTemp                As String
Dim wSATemp                 As String

wCFVTemp = "CFV_Temp"
wSATemp = "SA_Temp"

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wSATemp).Delete
Worksheets(wCFVTemp).Delete
Err.Clear

Application.DisplayAlerts = True
Sheets("Lookup").Activate
Range("AA1").Clear
Sheets("Pivot").Activate

End Sub


Sub WFM_Raw_Reports()

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

Sub WFM_Postprocess_Report()

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