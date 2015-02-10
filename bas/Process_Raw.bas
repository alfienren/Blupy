Attribute VB_Name = "Process_Raw"
Option Explicit

Sub Process_Raw_Reports()

Dim rFloodlightCell         As Range
Dim rSAData                 As Range
Dim rCFVData                As Range

Dim wCFVTemp                As String
Dim wSATemp                 As String
Dim wWorking                As String

With Application
    
    .ScreenUpdating = True
    .EnableEvents = False
    .Calculation = xlCalculationManual
    
End With

wCFVTemp = "CFV_Temp"
wSATemp = "SA_Temp"
wWorking = "working"

' Activate the Site Activity Sheet

Sheets("SA").Activate

With ActiveSheet
    
    ' Select the Site Activity data, excluding the top rows
    .Range("C1").Select
    Selection.End(xlDown).Select
    Range(Selection.End(xlToRight), Selection.End(xlToLeft)).Select
    
    ' Set reference to data to be used later
    Set rSAData = Range(Selection, Selection.End(xlDown).Offset(-1, 0))

End With

' If the SA_Temp sheet already exists, delete it

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wSATemp).Delete
Err.Clear

' Add worksheet named SA_Temp

Application.DisplayAlerts = True
Worksheets.Add.Name = wSATemp

' Activate the new sheet

Sheets("SA_Temp").Activate

With ActiveSheet

    ' Copy the Site Activity data referenced above into the worksheet starting
    ' at cell A1
    
    Cells.ClearContents
    rSAData.Copy
    .Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
End With

' Activate the raw CFV data tab

Sheets("CFV").Activate

With ActiveSheet

    ' As a simple check to make sure the data entered in the CFV tab is proper, search for the
    ' cell containing "Floodlight Attribution Type".
    
    Set rFloodlightCell = Cells.Find(What:="Floodlight Attribution Type", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        
    If rFloodlightCell Is Nothing Then
    
        ' If the cell is not found, send a message and exit the sub
        
        MsgBox "Make sure correct CFV data is entered in the tab"
        Exit Sub
        
    End If
    
    ' If the Floodlight Attribution Type cell is found, select the range in the CFV tab containing
    ' the data
    
    rFloodlightCell.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    ' Set a reference to the CFV data to be used later
    
    Set rCFVData = Range(Selection, Selection.End(xlDown).Offset(-1, 0))
    
End With

' If the CFV_Temp worksheet already exists, delete it

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wCFVTemp).Delete
Err.Clear

' Create a new worksheet named CFV_Temp

Application.DisplayAlerts = True
Worksheets.Add.Name = wCFVTemp
    
' Copy the CFV data referenced earlier

rCFVData.Copy

' Activate the CFV_Temp sheet

Sheets("CFV_Temp").Activate

With ActiveSheet
    
    ' Copy the CFV data into the CFV_Temp sheet starting at cell A1
    
    Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
                
End With

' If worksheet "working" already exists, delete it

Application.DisplayAlerts = False
On Error Resume Next
Worksheets(wWorking).Delete
Err.Clear

' Create new worksheet named "working"

Application.DisplayAlerts = True
Worksheets.Add.Name = wWorking

End Sub
