Attribute VB_Name = "TrafficSummaryData"
Sub TrafficSummary()

Dim rActualCell                 As Range
Dim rLastDateCell               As Range
Dim rDateRange                  As Range
Dim rTotalTrafficTypes          As Range
Dim rDirectLoadTypes            As Range
Dim rPerformanceTypes           As Range
Dim rPerformanceChannels        As Range
Dim rNonPerformanceTypes        As Range
Dim rNonPerformanceChannels     As Range
Dim i                           As Range
Dim j                           As Range
Dim raw                         As String

raw = "raw_data"

With Application

    .EnableEvents = False
    .ScreenUpdating = True
    .DisplayAlerts = False
    
End With

On Error Resume Next
Worksheets(raw).Delete
Err.Clear
Worksheets.Add.Name = raw

With Worksheets(raw)
    .Range("A1").Value = "Day of Week"
    .Range("B1").Value = "Date"
    .Range("C1").Value = "Traffic Type"
    .Range("D1").Value = "All Up Non-Deduped"
    .Range("E1").Value = "All Up Deduped"
    .Range("F1").Value = "Customer Non-Deduped"
    .Range("G1").Value = "Customer Deduped"
    .Range("H1").Value = "Prospect Non-Deduped"
    .Range("I1").Value = "Prospect Deduped"
    .Range("J1").Value = "Mobile Non-Deduped"
    .Range("K1").Value = "Mobile Deduped"
    .Range("L1").Value = "Non-Mobile Non-Deduped"
    .Range("M1").Value = "Non-Mobile Deduped"
End With

Sheets("Traffic Detail").Activate

Columns("C:G").Select
Set rActualCell = Selection.Find(What:="Actuals", After:=ActiveCell, LookIn:=xlFormulas, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)

Set rLastDateCell = rActualCell.Offset(1, 0).End(xlToRight).End(xlToRight).End(xlToRight)

' Dates and Day of Week
Set rDateRange = Range(rLastDateCell.Offset(-3, 0), rLastDateCell.Offset(-4, 0).End(xlToLeft))

' Total Traffic

Set rTotalTrafficTypes = Range(rActualCell.Offset(1, 0).End(xlToRight), _
    rActualCell.Offset(1, 0).End(xlToRight).End(xlDown))

rDateRange.Copy
    
Worksheets(raw).Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=True
    
Set j = Worksheets(raw).Range("E2")

For Each i In rTotalTrafficTypes
    
    Range(i.End(xlToRight), i.End(xlToRight).End(xlToRight)).Copy
    
    j.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    
    Set j = j.Offset(0, 2)
        
Next i

Sheets(raw).Activate

Range("B2", Range("B2").End(xlDown)).Offset(0, 1).Value = "Total"

' Direct Load Traffic

Sheets("Traffic Detail").Activate

Set rDirectLoadTypes = Range(rTotalTrafficTypes.End(xlDown).End(xlDown).Offset(0, 1), _
    rTotalTrafficTypes.End(xlDown).End(xlDown).Offset(0, 1).End(xlDown))

rDateRange.Copy
 
Worksheets(raw).Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=True

Sheets(raw).Activate

Set j = Worksheets(raw).Range("C1").End(xlDown).Offset(1, 0)

Range(j.Offset(0, -1), j.Offset(0, -1).End(xlDown)).Offset(0, 1).Value = "Total Direct Load"

Set j = j.Offset(0, 1)

For Each i In rDirectLoadTypes

    Range(i.End(xlToRight), i.End(xlToRight).End(xlToRight)).Copy
    j.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    
    Set j = j.Offset(0, 1)

Next i

' Total Traffic - Performance Marketing

Sheets("Traffic Detail").Activate

Set rPerformanceTypes = Range(rDirectLoadTypes.End(xlDown).End(xlDown), _
    rDirectLoadTypes.End(xlDown).End(xlDown).End(xlDown))
    
rDateRange.Copy

Worksheets(raw).Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=True

Sheets(raw).Activate

Set j = Worksheets(raw).Range("C1").End(xlDown).Offset(1, 0)

Range(j.Offset(0, -1), j.Offset(0, -1).End(xlDown)).Offset(0, 1).Value = "Total Performance Marketing"

Set j = j.Offset(0, 1)

For Each i In rPerformanceTypes

    Range(i.End(xlToRight), i.End(xlToRight).End(xlToRight)).Copy
    j.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    
    Set j = j.Offset(0, 1)

Next i

' Total Traffic - Performance Marketing - Channels

Sheets("Traffic Detail").Activate

Set rPerformanceChannels = Range(rPerformanceTypes.End(xlDown).End(xlDown), _
    rPerformanceTypes.End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).Offset(1, 0))
    
rDateRange.Copy

Worksheets(raw).Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
Sheets(raw).Activate

Set j = Worksheets(raw).Range("C1").End(xlDown).Offset(1, 1)

For Each i In rPerformanceChannels

    Range(i.End(xlToRight), i.End(xlToRight).End(xlToRight)).Copy
    j.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    
    Set j = j.Offset(0, 1)
    
    If IsEmpty(i) = True Then
    
        Set j = Worksheets(raw).Range("C1").End(xlDown).Offset(1, 0)
        Range(j.Offset(0, -1), j.Offset(0, -1).End(xlDown)).Offset(0, 1).Value = i.Offset(0, -3).End(xlUp).Value
        
        rDateRange.Copy

        Worksheets(raw).Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, _
            Operation:=xlNone, SkipBlanks:=False, Transpose:=True
            
        Set j = Worksheets(raw).Range("C1").End(xlDown).Offset(1, 1)
        
    End If
    
Next i

' Total Traffic - Non-Performance Marketing

Sheets("Traffic Detail").Activate

Set rNonPerformanceTypes = Range(rPerformanceChannels.End(xlToLeft).End(xlToLeft).End(xlToLeft).End(xlDown).End(xlToRight).End(xlToRight), _
    rPerformanceChannels.End(xlToLeft).End(xlToLeft).End(xlToLeft).End(xlDown).End(xlToRight).End(xlToRight).End(xlDown))

Sheets(raw).Activate

Set j = Worksheets(raw).Range("C1").End(xlDown).Offset(1, 0)

Range(j.Offset(0, -1), j.Offset(0, -1).End(xlDown)).Offset(0, 1).Value = "Total Non-Performance Marketing"

Set j = j.Offset(0, 1)

For Each i In rNonPerformanceTypes

    Range(i.End(xlToRight), i.End(xlToRight).End(xlToRight)).Copy
    j.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

    Set j = j.Offset(0, 1)

Next i

' Total Traffic - Non-Performance Marketing Channels

Sheets("Traffic Detail").Activate

Set rNonPerformanceChannels = Range(rNonPerformanceTypes.End(xlDown).End(xlDown), _
    rNonPerformanceTypes.End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).End(xlDown).Offset(1, 0))
    
rDateRange.Copy

Worksheets(raw).Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
Sheets(raw).Activate

Set j = Worksheets(raw).Range("C1").End(xlDown).Offset(1, 1)

For Each i In rNonPerformanceChannels

    Range(i.End(xlToRight), i.End(xlToRight).End(xlToRight)).Copy
    j.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    
    Set j = j.Offset(0, 1)
    
    If IsEmpty(i) = True Then
    
        Set j = Worksheets(raw).Range("C1").End(xlDown).Offset(1, 0)
        Range(j.Offset(0, -1), j.Offset(0, -1).End(xlDown)).Offset(0, 1).Value = i.Offset(0, -3).End(xlUp).Value
        
        rDateRange.Copy

        Worksheets(raw).Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, _
            Operation:=xlNone, SkipBlanks:=False, Transpose:=True
            
        Set j = Worksheets(raw).Range("C1").End(xlDown).Offset(1, 1)
        
    End If
    
Next i

Range(Range("C1").End(xlDown).Offset(1, 0), Range("C1").End(xlDown).Offset(1, 0).End(xlDown)).EntireRow.Delete Shift:=xlUp
Worksheets(raw).Columns("B:B").NumberFormat = "m/d/yyyy"

End Sub
