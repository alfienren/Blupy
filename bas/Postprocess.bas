Attribute VB_Name = "Postprocess"
Option Explicit

Sub Postprocess_Report()

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

End Sub
