Attribute VB_Name = "Tools_PivotGenerate"
Option Explicit

Sub GeneratePivot()

Dim rData       As Range
Dim wPivot      As Worksheet
Dim oPivot      As Object
Dim sPivot      As String

Sheets("data").Activate

Set rData = Range("A1").CurrentRegion

sPivot = "pivot"

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(sPivot).Delete

Err.Clear

Application.DisplayAlerts = True
Worksheets.Add.Name = sPivot

ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    rData, Version:=xlPivotTableVersion15).CreatePivotTable _
    TableDestination:="pivot!R3C1", TableName:="Pivot", DefaultVersion _
    :=xlPivotTableVersion15
    
Set wPivot = ActiveSheet
Set oPivot = wPivot.PivotTables("Pivot")

oPivot.CalculatedFields.Add "Traffic Yield", _
    "=IFERROR('Traffic Actions'/Impressions,0)", True

oPivot.CalculatedFields.Add "Video Completion Rate", _
    "=IFERROR('Video Views'/'Video Completions',0)", True

oPivot.CalculatedFields.Add "Cost Per Traffic Action", _
    "=IFERROR('NTC Media Cost'/'Traffic Actions',0)", True
    
With oPivot.PivotFields("Campaign")
    .Orientation = xlRowField
    .Position = 1
End With

With oPivot.PivotFields("Site")
    .Orientation = xlRowField
    .Position = 2
End With

With oPivot.PivotFields("Week")
    .Orientation = xlRowField
    .Position = 3
End With

With oPivot
    .AddDataField oPivot.PivotFields("Impressions"), "Sum of Impressions", xlSum
    .AddDataField oPivot.PivotFields("NTC Media Cost"), "Sum of NTC Media Cost", xlSum
    .AddDataField oPivot.PivotFields("Traffic Actions"), "Sum of Traffic Actions", xlSum
    .AddDataField oPivot.PivotFields("Traffic Yield"), "Sum of Traffic Yield", xlSum
    .AddDataField oPivot.PivotFields("Cost Per Traffic Action"), "Sum of Cost Per Traffic Actions", xlSum
End With

With oPivot
    .PivotFields("Campaign").RepeatLabels = True
    .PivotFields("Site").RepeatLabels = True
    .PivotFields("Week").RepeatLabels = True
End With

With oPivot
    .InGridDropZones = True
    .RowAxisLayout xlTabularRow
End With

End Sub
