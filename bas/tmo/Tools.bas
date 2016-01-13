Attribute VB_Name = "Tools"
Option Private Module
Option Explicit

Sub DataSplit()

Dim column         As String

column = Application.InputBox(Prompt:="Enter desired column name to split data.", Title:="Run Split Data Macro", Type:=2)
              
If column <> vbNullString Then

    Sheets("Lookup").Activate
    Range("AA1").Value = ThisWorkbook.FullName
    Range("AB1").Value = column
    
    Call Python_Split_Data
    
Else

    Exit Sub
    
End If

End Sub

Sub MergeData()

Dim sData           As String
Dim sDataSheet      As String

sData = Application.InputBox(Prompt:="Enter path of workbook containing data to merge", Title:="Merge Data", Type:=2)

If Right(sData, 4) <> ".csv" Then

    sDataSheet = Application.InputBox(Prompt:="Enter name of tab containing data", Title:="Worksheet Name", Type:=2)
    Range("AC2").Value = sDataSheet

End If

Sheets("Lookup").Activate
Range("AA2").Value = ThisWorkbook.FullName
Range("AB2").Value = sData

If sData <> vbNullString Then

    Call Python_Merge_Data
    Range("AA2:AC2").ClearContents

Else

    Exit Sub
    
End If

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

