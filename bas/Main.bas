Attribute VB_Name = "Main"
Option Explicit

Sub DFA_Reporting()

Dim rDDR    As Range

Call Process_Raw_Reports

Call Select_Feed_File

Call Python_Weekly_Reporting

Call Top_15_Devices
Call Postprocess_Report

End Sub

Sub Compress_Data()

Dim wOriginal   As Workbook
Dim wOutput     As Workbook
Dim wpath       As String

Set wOriginal = ThisWorkbook

Sheets("Tools").Activate
Range("ZZ1").Value = ThisWorkbook.FullName

Call Python_Compress_Data

Set wOutput = ActiveWorkbook

Call CopyModule(wOriginal, "Pivot_Generate", wOutput)

With ActiveWorkbook

    Call GeneratePivot

End With

ActiveWorkbook.Save
ActiveWorkbook.Close

End Sub
