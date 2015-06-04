Attribute VB_Name = "Tools"
Option Explicit

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
