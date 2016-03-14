Attribute VB_Name = "Campaign_Pacing"
Option Explicit

Sub Flat_Rate_Placement_Report()

Dim wFlatRates          As String

wFlatRates = "Flat_Rate"

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wFlatRates).Delete
Err.Clear

Application.DisplayAlerts = True
Worksheets.Add.Name = wFlatRates

Call Select_Planned_Media_Report

Call Python_FlatRates

Sheets("Action_Reference").Activate
Range("AE1").Clear

Sheets("Flat_Rate").Activate

End Sub

Sub Campaign_Pacing_Report()

Call Select_Trafficking_Campaign_Master_File

Call Python_Pacing_Report

Sheets("Action_Reference").Activate
Range("AE1").Clear

Sheets("data").Activate

End Sub
