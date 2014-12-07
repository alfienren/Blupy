Attribute VB_Name = "Process_Raw_Reports"
Option Explicit

Sub ProcessRawDFA()

Dim First                   As Long
Dim Last                    As Long
Dim Lrow                    As Long

Dim rUniqueID               As Range
Dim rFloodlightCell         As Range

Dim rSAData                 As Range
Dim rUniqueIDSA             As Range

Dim rUniqueIDCFV            As Range
Dim rCFVData                As Range
Dim rCFVCell                As Range

Dim cell                    As Range
Dim cell2                   As Range
Dim var                     As Range

With Application
    
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
    
End With

Sheets("SA").Activate

With ActiveSheet

    Set rUniqueID = Cells.Find(What:="Unique ID", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        
    If rUniqueID Is Nothing Then
    
    Else
    
        rUniqueID.Columns.Delete Shift:=xlToLeft
        
    End If
    
    .Range("C1").Select
    Selection.End(xlDown).Select
    Range(Selection.End(xlToRight), Selection.End(xlToLeft)).Select
    
    Set rSAData = Range(Selection, Selection.End(xlDown))

End With

Sheets("working").Activate

With ActiveSheet

    Cells.ClearContents
    rSAData.Copy
    .Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    .Range("B1").End(xlDown).EntireRow.Delete
    
    .Range("A1").Value = "Unique ID"
    .Range("A2").Select
    
    Range(Selection, Selection.Offset(rSAData.Rows.Count - 3, 0)).Select
    
    Selection.Formula = "=RC[1]&RC[2]&RC[3]&RC[9]&RC[12]"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    Set rUniqueIDSA = Selection
    
End With

Sheets("CFV").Activate

With ActiveSheet

    Set rFloodlightCell = Cells.Find(What:="Floodlight Attribution Type", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)

    Set rUniqueID = Cells.Find(What:="Unique ID", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        
    If rUniqueID Is Nothing Then
    
    Else
    
        rUniqueID.EntireColumn.Delete Shift:=xlToLeft
        
    End If
    
    Range(rFloodlightCell, rFloodlightCell.End(xlToRight)).Select
    Set rCFVData = Range(Selection, Selection.End(xlDown))
    
    .Range("C1").Select
    Selection.End(xlDown).Select
    
    Range(Selection, Selection.End(xlDown)).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Selection.Value = "Unique ID"
    
    Selection.Offset(1, -1).Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 1).Select
    
    Selection.Formula = "=RC[-2]&RC[-1]&RC[1]&RC[7]&RC[9]"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    Set rUniqueIDCFV = Selection
    
End With

Set rCFVCell = rUniqueIDCFV.Cells(1)

Sheets("Lookup").Activate
With ActiveSheet

    rUniqueIDSA.Copy
    .Range("AA1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    .Range("AA2").Select
    Set rUniqueIDSA = Range(Selection, Selection.End(xlDown))

    rUniqueIDCFV.Copy
    .Range("AB2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    .Range("AB1").Select
    Set rUniqueIDCFV = Range(Selection, Selection.End(xlDown))
        
    rCFVData.Copy
    .Range("AC1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End With

Sheets("working").Activate

With ActiveSheet

    With .Range("A1").End(xlToRight)
    
        .Offset(0, 1).Value = "Floodlight Attribution Type"
        
        .Offset(1, 1).Select
        
        For Each cell In rUniqueIDSA
                            
            var = Application.WorksheetFunction.Match(cell, rUniqueIDCFV, 0)
                
            If Not IsError(var) Then
                
                Selection.Value = var.Offset(0, 2)
                    
            Else
                
                Selection.Value = 0
                    
            End If
                
        Selection.Offset(1, 0).Select
            
        Next cell
        
        .Offset(0, 2).Value = "Activity"
        .Offset(0, 3).Value = "Order Number"
        .Offset(0, 4).Value = "Plan (string)"
        .Offset(0, 5).Value = "Device (string)"
        .Offset(0, 6).Value = "Service (string)"
        .Offset(0, 7).Value = "Accessory (string)"
        .Offset(0, 8).Value = "Transaction Count"

    End With
    
End With

End Sub
