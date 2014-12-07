Attribute VB_Name = "DataMunge"
Sub MungeData()

Dim First As Long
Dim Last As Long
Dim Lrow As Long
Dim CalcMode As Long
Dim ViewMode As Long
Dim Clicks As Range
Dim Floodlight As Range
Dim rng As Range
Dim cell As Range
Dim SumRange As Range
Dim ActionRange As Range
Dim SearchRange As Range
Dim FindWhat As Variant
Dim FoundCells As Range
Dim FoundCell As Range
Dim ERange As Range
Dim FRange As Range
Dim UniqueID As Range

'Delete old data in sheets if present
Sheets("working_revised").Activate
    Cells.Select
    Selection.ClearContents
    
Sheets("SA").Activate

Set SearchRange = Range("A:ZZ")
FindWhat = "Unique ID"

Set FoundCells = FindAll(SearchRange:=SearchRange, _
                        FindWhat:=FindWhat, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByColumns, _
                        MatchCase:=False, _
                        BeginsWith:=vbNullString, _
                        EndsWith:=vbNullString, _
                        BeginEndCompare:=vbTextCompare)
                        
If FoundCells Is Nothing Then

Else
    FoundCells.Columns.Select
    Selection.Delete Shift:=xlToLeft
End If

Sheets("pivot_data").Activate
    Range("D3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("G3:I3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Columns("AF:EZ").Select
    Selection.ClearContents

Sheets("CFV").Activate

Set SearchRange = Range("A:ZZ")
FindWhat = "Unique ID"

Set FoundCells = FindAll(SearchRange:=SearchRange, _
                        FindWhat:=FindWhat, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByColumns, _
                        MatchCase:=False, _
                        BeginsWith:=vbNullString, _
                        EndsWith:=vbNullString, _
                        BeginEndCompare:=vbTextCompare)
                        
If FoundCells Is Nothing Then

Else
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
End If

Sheets("SA").Activate

Set SearchRange = Range("A:A")
FindWhat = "Report Fields"

Set FoundCells = FindAll(SearchRange:=SearchRange, _
                        FindWhat:=FindWhat, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByColumns, _
                        MatchCase:=False, _
                        BeginsWith:=vbNullString, _
                        EndsWith:=vbNullString, _
                        BeginEndCompare:=vbTextCompare)
                        
If FoundCells Is Nothing Then

Else
    FoundCells.ClearContents
End If

Range("A1").Select
Selection.End(xlDown).Select
Selection.End(xlDown).Select
Selection.End(xlDown).Select

If Selection.Value = "Campaign" Then
    Selection.CurrentRegion.Select
Else
    Set SearchRange = Range("A14:T1000")
    FindWhat = "Campaign"

    Set FoundCells = FindAll(SearchRange:=SearchRange, _
                            FindWhat:=FindWhat, _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByColumns, _
                            MatchCase:=False, _
                            BeginsWith:=vbNullString, _
                            EndsWith:=vbNullString, _
                            BeginEndCompare:=vbTextCompare)
                            
    FoundCells.CurrentRegion.Select
End If


'Delete Grand Total Row from SA sheet

With Application
    CalcMode = .Calculation
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = False

End With

With ActiveSheet
    First = .UsedRange.Cells(1).row
    Last = .UsedRange.Rows(.UsedRange.Rows.Count).row
    
    For Lrow = Last To First Step -1
        With .Cells(Lrow, "A")
            If Not IsError(.Value) Then
                If .Value = "Grand Total:" Then .EntireRow.Delete
            End If
        End With
    Next Lrow
End With


Set SearchRange = Range("A:ZZ")
FindWhat = "Unique ID"

Set UniqueID = FindAll(SearchRange:=SearchRange, _
                        FindWhat:=FindWhat, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByColumns, _
                        MatchCase:=False, _
                        BeginsWith:=vbNullString, _
                        EndsWith:=vbNullString, _
                        BeginEndCompare:=vbTextCompare)
                        
If UniqueID Is Nothing Then
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select

    Selection.End(xlToRight).Select
    Selection.Offset(0, 1).Select
    Selection.Value = "Unique ID"

    Selection.Offset(1, -1).Select
    Range(Selection, Selection.End(xlDown)).Offset(0, 1).Select

    Selection.Formula = "=RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-1]"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
End If

Selection.CurrentRegion.Copy

Sheets("working_revised").Activate
Range("A1").Select
ActiveSheet.Paste
        
Sheets("CFV").Activate

Set SearchRange = Range("A:A")
FindWhat = "Report Fields"

Set FoundCells = FindAll(SearchRange:=SearchRange, _
                        FindWhat:=FindWhat, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByColumns, _
                        MatchCase:=False, _
                        BeginsWith:=vbNullString, _
                        EndsWith:=vbNullString, _
                        BeginEndCompare:=vbTextCompare)
                        
If FoundCells Is Nothing Then

Else
    FoundCells.ClearContents
End If

Range("A1").Select
Selection.End(xlDown).Select
Selection.End(xlDown).Select
Selection.End(xlDown).Select

If Selection.Value = "Campaign" Then
    Selection.CurrentRegion.Select
Else
    Set SearchRange = Range("A10:T1000")
    FindWhat = "Campaign"

    Set FoundCells = FindAll(SearchRange:=SearchRange, _
                            FindWhat:=FindWhat, _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByColumns, _
                            MatchCase:=False, _
                            BeginsWith:=vbNullString, _
                            EndsWith:=vbNullString, _
                            BeginEndCompare:=vbTextCompare)
                            
    FoundCells.CurrentRegion.Select
End If

Set SearchRange = Range("A:ZZ")
FindWhat = "Unique ID"

Set UniqueID = FindAll(SearchRange:=SearchRange, _
                        FindWhat:=FindWhat, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByColumns, _
                        MatchCase:=False, _
                        BeginsWith:=vbNullString, _
                        EndsWith:=vbNullString, _
                        BeginEndCompare:=vbTextCompare)
                        
If UniqueID Is Nothing Then
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                            
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
                            
    Selection.Offset(1, 0).Select
    Selection.End(xlToRight).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Offset(0, 1).Select
    Selection.Formula = "=RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-1]"
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
        
    Selection.End(xlToRight).Select
    Selection.Offset(0, 1).Select
    Selection.Value = "Unique ID"
End If

'Copy Floodlight headers in CFV sheet to working data tab
Set SearchRange = Range("A1:Z1000")
FindWhat = "Floodlight Attribution Type"

Set FoundCells = FindAll(SearchRange:=SearchRange, _
                        FindWhat:=FindWhat, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByColumns, _
                        MatchCase:=False, _
                        BeginsWith:=vbNullString, _
                        EndsWith:=vbNullString, _
                        BeginEndCompare:=vbTextCompare)
                        
FoundCells.Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

Sheets("working_revised").Activate
Range("A1").Select
Selection.End(xlToRight).Select
Selection.Offset(0, 1).Select
ActiveSheet.Paste

'Clean up URLs
    Columns("E:E").Select
    Selection.Replace What:="%3F%26csdids%3DADV_DS_%epid!_%eaid!_%ecid!_%eadv!", _
        Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
    
    Selection.Replace What:= _
        "analytics.bluekai.com/site/15991?phint=event%3Dclick&phint=aid%3D%eadv!&phint=pid%3D%epid!&phint=cid%3D%ebuy!&phint=crid%3D%ecid!&done=http%3A%2F%2F" _
        , Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
    Selection.Replace What:="%2F", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Selection.Replace What:="http://analytics.bluekai.com/site/15991?phint=event%3Dclick&phint=aid%3D%eadv!&phint=pid%3D%epid!&phint=cid%3D%ebuy!&phint=crid%3D%ecid!&done=", _
        Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
    Selection.Replace What:="?cid=dis-", _
        Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
'Merge Data from CFV sheet into working data tab
Range("A2").Select
Selection.End(xlToRight).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Offset(0, 1).Select

'Floodlight Attribution Type
Selection.Formula = "=IFERROR(VLOOKUP(INDEX(R,0,MATCH(""Unique ID"",R1,0)),CFV!C9:C20,5,FALSE),0)"

'Activity
Selection.Offset(0, 1).Select
Selection.Formula = "=IFERROR(VLOOKUP(INDEX(R,0,MATCH(""Unique ID"",R1,0)),CFV!C9:C20,6,FALSE),0)"

'Order Number
Selection.Offset(0, 1).Select
Selection.Formula = "=IFERROR(VLOOKUP(INDEX(R,0,MATCH(""Unique ID"",R1,0)),CFV!C9:C20,7,FALSE),0)"

'Plan
Selection.Offset(0, 1).Select
Selection.Formula = "=IFERROR(VLOOKUP(INDEX(R,0,MATCH(""Unique ID"",R1,0)),CFV!C9:C20,8,FALSE),0)"

'Device
Selection.Offset(0, 1).Select
Selection.Formula = "=IFERROR(VLOOKUP(INDEX(R,0,MATCH(""Unique ID"",R1,0)),CFV!C9:C20,9,FALSE),0)"

'Service
Selection.Offset(0, 1).Select
Selection.Formula = "=IFERROR(VLOOKUP(INDEX(R,0,MATCH(""Unique ID"",R1,0)),CFV!C9:C20,10,FALSE),0)"

'Accessory
Selection.Offset(0, 1).Select
Selection.Formula = "=IFERROR(VLOOKUP(INDEX(R,0,MATCH(""Unique ID"",R1,0)),CFV!C9:C20,11,FALSE),0)"

'Transaction Count
Selection.Offset(0, 1).Select
Selection.Formula = "=IFERROR(SUMPRODUCT(--(CFV!C9=INDEX(R,0,MATCH(""Unique ID"",R1,0))),CFV!C20),0)"

Range("A1").CurrentRegion.Select
Selection.Copy

Sheets("pivot_data").Activate
Range("AF2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Range("A3").Select
Selection.End(xlToRight).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Offset(0, 1).Select

'Category
Selection.Formula = "=VLOOKUP(LOOKUP(9.9999E+307,SEARCH(Lookup!R2C1:R7C1,pivot_data!RC[40]),Lookup!R2C1:R7C1),Lookup!C1:C2,2,FALSE)"

'F Tag
Range("F3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Formula = "=VLOOKUP(RC[-1],Lookup!C[-2]:C[-1],2,FALSE)"

'Message Categorization
Range("A3").Select
Selection.End(xlToRight).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Offset(0, 1).Select

'Message Bucket
Selection.Formula = "=TRIM(MID(SUBSTITUTE(SUBSTITUTE(RC[36],""Creative Type: "",""""),""_"",REPT("" "",99)),COLUMN(R[-2]C[-6])*99-98,99))"

'Message Category
Selection.Offset(0, 1).Select
Selection.Formula = "=TRIM(MID(SUBSTITUTE(SUBSTITUTE(RC[35],""Creative Type: "",""""),""_"",REPT("" "",99)),COLUMN(R[-2]C[-6])*99-98,99))"

'Message Offer
Selection.Offset(0, 1).Select
Selection.Formula = "=TRIM(MID(SUBSTITUTE(SUBSTITUTE(RC[34],""Creative Type: "",""""),""_"",REPT("" "",99)),COLUMN(R[-2]C[-6])*99-98,99))"

'Find F tag conversions
Range("AF2").Select
Range(Selection, Selection.End(xlToRight)).Select

Set Clicks = Columns("AF:ZZ").Find(What:="Clicks", After:=ActiveCell, LookIn:=xlValues, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Offset(0, 1)
    
Set Floodlight = Columns("AF:ZZ").Find(What:="Floodlight Attribution Type", After:=ActiveCell, _
    LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Offset(0, -1)

Range("Z3").Select
    
Set rng = Range(Selection, Selection.End(xlDown))
Set SumRange = Range(Clicks.Offset(1, 0), Floodlight.Offset(1, 0))
Set ActionRange = Range("V3:Y3")

Sheets("Lookup").Activate

Set ERange = Range("AZ1")
    
For Each cell In rng
    
    cell = Application.WorksheetFunction.Sum(SumRange) - Application.WorksheetFunction.Sum(ActionRange) - Application.WorksheetFunction.Sum(ERange)
    Set SumRange = SumRange.Offset(1, 0)
    Set ActionRange = ActionRange.Offset(1, 0)
    Set ERange = ERange.Offset(1, 0)
        
Next cell

Sheets("pivot_data").Activate
Set SearchRange = Range("AF2:ZZ2")
FindWhat = "View-through"

Set FoundCells = FindAll(SearchRange:=SearchRange, _
                        FindWhat:=FindWhat, _
                        LookIn:=xlValues, _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByColumns, _
                        MatchCase:=False, _
                        BeginsWith:=vbNullString, _
                        EndsWith:=vbNullString, _
                        BeginEndCompare:=vbTextCompare)
                        
Range("AE3").Select
Set rng = Range(Selection, Selection.End(xlDown))

For Each cell In rng

    Set FoundCells = FoundCells.Offset(1, 0)
    cell = Application.WorksheetFunction.Sum(FoundCells)
    
Next cell

Set SearchRange = Range("AF2:ZZ2")
FindWhat = "Click-through"

Set FoundCells = FindAll(SearchRange:=SearchRange, _
                        FindWhat:=FindWhat, _
                        LookIn:=xlValues, _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByColumns, _
                        MatchCase:=False, _
                        BeginsWith:=vbNullString, _
                        EndsWith:=vbNullString, _
                        BeginEndCompare:=vbTextCompare)
                        
Range("AD3").Select
Set rng = Range(Selection, Selection.End(xlDown))


For Each cell In rng

    Set FoundCells = FoundCells.Offset(1, 0)
    cell = Application.WorksheetFunction.Sum(FoundCells)
    
Next cell

'Formatting
    Columns("F:DQ").EntireColumn.AutoFit

    Columns("AG:AG").Select
    Selection.NumberFormat = "m/d/yyyy"
    
End Sub

Function FindAll(SearchRange As Range, _
                FindWhat As Variant, _
                Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlPart, _
                Optional SearchOrder As XlSearchOrder = xlByRows, _
                Optional MatchCase As Boolean = False, _
                Optional BeginsWith As String = vbNullString, _
                Optional EndsWith As String = vbNullString, _
                Optional BeginEndCompare As VbCompareMethod = vbTextCompare) As Range
                
Dim FoundCell As Range
Dim FirstFound As Range
Dim LastCell As Range
Dim ResultRange As Range
Dim XLookAt As XlLookAt
Dim Include As Boolean
Dim CompMode As VbCompareMethod
Dim Area As Range
Dim MaxRow As Long
Dim MaxCol As Long
Dim BeginB As Boolean
Dim EndB As Boolean

CompMode = BeginEndCompare
If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then
    XLookAt = xlPart
Else
    XLookAt = LookAt
End If

' this loop in Areas is to find the last cell
' of all the areas. That is, the cell whose row
' and column are greater than or equal to any cell
' in any Area.
For Each Area In SearchRange.Areas
    With Area
        If .Cells(.Cells.Count).row > MaxRow Then
            MaxRow = .Cells(.Cells.Count).row
        End If
        If .Cells(.Cells.Count).Column > MaxCol Then
            MaxCol = .Cells(.Cells.Count).Column
        End If
    End With
Next Area
Set LastCell = SearchRange.Worksheet.Cells(MaxRow, MaxCol)


'On Error Resume Next
On Error GoTo 0
Set FoundCell = SearchRange.Find(What:=FindWhat, _
        After:=LastCell, _
        LookIn:=LookIn, _
        LookAt:=XLookAt, _
        SearchOrder:=SearchOrder, _
        MatchCase:=MatchCase)

If Not FoundCell Is Nothing Then
    Set FirstFound = FoundCell
    'Set ResultRange = FoundCell
    'Set FoundCell = SearchRange.FindNext(after:=FoundCell)
    Do Until False ' Loop forever. We'll "Exit Do" when necessary.
        Include = False
        If BeginsWith = vbNullString And EndsWith = vbNullString Then
            Include = True
        Else
            If BeginsWith <> vbNullString Then
                If StrComp(Left(FoundCell.Text, Len(BeginsWith)), BeginsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
            If EndsWith <> vbNullString Then
                If StrComp(Right(FoundCell.Text, Len(EndsWith)), EndsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
        End If
        If Include = True Then
            If ResultRange Is Nothing Then
                Set ResultRange = FoundCell
            Else
                Set ResultRange = Application.Union(ResultRange, FoundCell)
            End If
        End If
        Set FoundCell = SearchRange.FindNext(After:=FoundCell)
        If (FoundCell Is Nothing) Then
            Exit Do
        End If
        If (FoundCell.Address = FirstFound.Address) Then
            Exit Do
        End If

    Loop
End If
    
Set FindAll = ResultRange

End Function
