Attribute VB_Name = "DDR_Top15_Devices"
Option Explicit

Sub DDR_Top_15_Devices()

Dim wSummary                As String
Dim wDDR                    As String
Dim TempFilePath            As String
Dim TempFileName            As String
Dim FileExtStr              As String

Dim vFileName               As Variant
Dim wOriginal               As Workbook
Dim wDeviceReport           As Workbook
Dim lFileFormat             As Long
Dim ws                      As Worksheet

With Application
    
    .ScreenUpdating = True
    .EnableEvents = False
    .Calculation = xlCalculationManual
    
End With

wSummary = "Summary"
wDDR = "DDR"

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wSummary).Delete
Worksheets(wDDR).Delete
Err.Clear

Application.DisplayAlerts = True
Worksheets.Add.Name = wSummary
Worksheets.Add.Name = wDDR

Call Python_DDR_Top_Devices

Sheets("Summary").Activate

Range("A1").Value = "Rank"
Range("B1").Value = "Devices"
Range("C1").Value = "Devices Count"
Range("D1").Value = "Device Name"

Range("A1:E1").Font.Bold = True

With Range("A1:E1").Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark2
    .TintAndShade = -9.99786370433668E-02
    .PatternTintAndShade = 0
End With

TempFilePath = Application.ActiveWorkbook.Path & "\"
TempFileName = "Top 15 Devices Report" & Format(Date, "mmddyyyy") & ".xlsx"

Set wDeviceReport = Workbooks.Add

ThisWorkbook.Sheets("Summary").Copy Before:=wDeviceReport.Sheets(1)

Application.DisplayAlerts = False
wDeviceReport.Sheets("Sheet1").Delete
Application.DisplayAlerts = True

Columns("A:E").EntireColumn.AutoFit
Columns("B:B").NumberFormat = "0"

With Range("A1").CurrentRegion
    
    .Copy
    .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End With

Range("A1").Select

wDeviceReport.SaveAs FileName:=TempFilePath & TempFileName

End Sub
