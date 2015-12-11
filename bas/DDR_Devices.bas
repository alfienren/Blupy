Attribute VB_Name = "DDR_Devices"
Option Private Module
Option Explicit

Sub Top_15_Devices()

Dim TempFilePath            As String
Dim TempFileName            As String
Dim FileExtStr              As String
Dim wOriginal               As Workbook

Dim vFileName               As Variant
Dim wDeviceReport           As Workbook
Dim lFileFormat             As Long
Dim ws                      As Worksheet

Set wOriginal = ActiveWorkbook

With Application
    
    .ScreenUpdating = True
    .EnableEvents = False
    .Calculation = xlCalculationManual
    
End With

Sheets("Summary").Activate

Range("A1").Value = "Rank"
Range("B1").Value = "Devices"
Range("C1").Value = "Devices Count"
Range("D1").Value = "Device Name"
Range("H1").Value = "Rank"
Range("I1").Value = "Plans"
Range("J1").Value = "Plans Count"

Range("A1:D1").Font.Bold = True

With Range("A1:D1").Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark2
    .TintAndShade = -9.99786370433668E-02
    .PatternTintAndShade = 0
End With

Range("H1:J1").Font.Bold = True

With Range("H1:J1").Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark2
    .TintAndShade = -9.99786370433668E-02
    .PatternTintAndShade = 0
End With

TempFilePath = Application.ActiveWorkbook.path & "\"
TempFileName = "Top 15 Devices Report" & Format(Date, "mmddyyyy") & ".xlsx"

Set wDeviceReport = Workbooks.Add

ThisWorkbook.Sheets("Summary").Copy Before:=wDeviceReport.Sheets(1)

Application.DisplayAlerts = False
wDeviceReport.Sheets("Sheet1").Delete
Application.DisplayAlerts = True

Columns("A:F").EntireColumn.AutoFit
Columns("H:J").EntireColumn.AutoFit
Columns("B:B").NumberFormat = "0"

With Range("A1").CurrentRegion
    
    .Copy
    .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End With

Range("A1").Select

wDeviceReport.SaveAs FileName:=TempFilePath & TempFileName

wOriginal.Sheets("Summary").Activate

End Sub

