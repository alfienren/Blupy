Attribute VB_Name = "Main"
Option Explicit

Sub Metro_Reporting()

Call Process_Raw_Reports

Call Python_Weekly_Reporting

Call Postprocess_Report

End Sub



