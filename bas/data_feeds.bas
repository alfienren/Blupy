Attribute VB_Name = "data_feeds"
Sub TMO_Cost_Feed()

Dim rDR     As Range
Dim rBrand  As Range
Dim rPath   As Variant

Sheets("data").Activate

Columns("E:E").Select

Set rDR = Selection.Find(What:="DR", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    False, SearchFormat:=False)

Set rBrand = Selection.Find(What:="Brand Remessaging", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False)

If rDR Is Nothing Then

    Set sFile = Application.FileDialog(msoFileDialogFilePicker)
    
    With sFile
        .Title = "Choose DR Pivot"
        .AllowMultiSelect = False
        
        If .Show <> -1 Then
            Exit Sub
        End If
        
        FileSelected = .SelectedItems(1)
    
    End With
    
End If

Sheets("Action_Reference").Activate

Range("AA1").Value = ThisWorkbook.FullName
Range("AC1").Value = FileSelected

Call Python_eBay_CostFeed

Range("AA1").Clear
Range("AC1").Clear

Sheets("data").Activate

End Sub

Sub Select_Feed_File()

Set sFile = Application.FileDialog(msoFileDialogFilePicker)

With sFile

    .Title = "Select downloaded device feed text file"
    
    .AllowMultiSelect = False
    
    If .Show <> -1 Then
        Exit Sub
    End If
    
    FileSelected = .SelectedItems(1)
    
End With

Sheets("Action_Reference").Activate
Range("AE1").Value = FileSelected

End Sub
