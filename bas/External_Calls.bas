Attribute VB_Name = "External_Calls"
Option Explicit

Sub Python_Calls()

' Go to the Lookup tab and copy in the path and name of the Excel sheet so the
' Python module can reference the workbook properly

Sheets("Lookup").Activate
Range("G1").Value = ActiveWorkbook.FullName

' Run the Python module

RunPython ("import weekly_reporting; weekly_reporting.dfa_reporting()")

End Sub
