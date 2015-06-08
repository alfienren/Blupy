Attribute VB_Name = "Tools_SplitData"
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
