Attribute VB_Name = "Tools_Merge"
Option Explicit

Sub MergeData()

Dim sData           As String
Dim sDataSheet      As String

sData = Application.InputBox(Prompt:="Enter path of workbook containing data to merge", Title:="Merge Data", Type:=2)

sDataSheet = Application.InputBox(Prompt:="Enter name of tab containing data", Title:="Worksheet Name", Type:=2)

Sheets("Lookup").Activate
Range("AA2").Value = ThisWorkbook.FullName
Range("AB2").Value = sData
Range("AC2").Value = sDataSheet

If sData <> vbNullString Then

    Call Python_Merge_Data
    Range("AA2:AC2").ClearContents

Else

    Exit Sub
    
End If

End Sub
