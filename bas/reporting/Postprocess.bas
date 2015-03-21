Attribute VB_Name = "Postprocess"
Option Explicit

Sub Postprocess_Report()

Dim wCFVTemp                As String
Dim wSATemp                 As String
Dim wWorking                As String
Dim wSummary                As String
Dim wDDR                    As String

wCFVTemp = "CFV_Temp"
wSATemp = "SA_Temp"
wWorking = "working"
wSummary = "Summary"
wDDR = "DDR"

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wSATemp).Delete
Worksheets(wCFVTemp).Delete
Worksheets(wWorking).Delete
Worksheets(wSummary).Delete
Worksheets(wDDR).Delete
Err.Clear

Application.DisplayAlerts = True

Sheets("Lookup").Activate
Range("G1").ClearContents

Sheets("Pivot").Activate

End Sub
