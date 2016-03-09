Attribute VB_Name = "Main"
Option Explicit

' DDR currently being reported separately

Sub DFA_Reporting()

Dim rDDR    As Range

Call Process_Raw_Reports

Call Select_Feed_File

Call Python_Weekly_Reporting

'Call Top_15_Devices
Call Postprocess_Report

End Sub

Sub Create_Traffic_Master_Sheet()

Call Select_Campaign_Trafficking_Reports_Folder

Call Python_Build_Trafficking_Master_Sheet

Sheets("Action_Reference").Activate
Range("AE1").Clear

Sheets("data").Activate

End Sub
