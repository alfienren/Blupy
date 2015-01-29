Attribute VB_Name = "External_Calls"
Option Explicit

Sub Python_Calls()

Sheets("Lookup").Activate
Range("G1").Value = ActiveWorkbook.FullName

RunPython ("import weekly_reporting; weekly_reporting.dfa_reporting()")

End Sub
