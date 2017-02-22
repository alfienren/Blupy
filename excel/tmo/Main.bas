Attribute VB_Name = "Main"
Option Explicit

Sub Weekly_Reporting()

Dim rDDR    As Range

Call Process_Raw_Reports

Call Select_Feed_File

Call TMO_Weekly_Reporting

Call Postprocess_Report

End Sub

Sub Pacing_Report()

Call Select_Feed_File

Call TMO_Pacing_Report

End Sub

Sub Flat_Rate_Report()

Call Select_Feed_File

Call TMO_FlatRates

End Sub
