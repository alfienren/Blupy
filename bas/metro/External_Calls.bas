Attribute VB_Name = "External_Calls"
Option Explicit

Sub Python_Weekly_Reporting()

Sheets("Lookup").Activate
Range("AA1").Value = ActiveWorkbook.FullName

RunPython ("import metro_weekly; ")

End Sub
