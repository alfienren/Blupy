Attribute VB_Name = "dfa_Post_Process"
Option Private Module
Option Explicit

Sub Postprocess_Report()

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
