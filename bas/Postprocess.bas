Attribute VB_Name = "Postprocess"
Option Explicit

Sub Postprocess_Report()

Dim wCFVTemp                As String
Dim wSATemp                 As String
Dim wWorking                As String

wCFVTemp = "CFV_Temp"
wSATemp = "SA_Temp"
wWorking = "working"

' Now that the routine is complete, delete the three temporary tabs
' SA_Temp, CFV_Temp and working

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wSATemp).Delete
Worksheets(wCFVTemp).Delete
Worksheets(wWorking).Delete
Err.Clear

Application.DisplayAlerts = True

' Delete the path and name reference from the Lookup tab

Sheets("Lookup").Activate
Range("G1").ClearContents

' Go to the Pivot tab

Sheets("Pivot").Activate

End Sub
