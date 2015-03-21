Attribute VB_Name = "Split_Sheets"
Option Explicit

Sub SplitSheets()

Dim Filestr         As String
Dim FileFormat      As Long
Dim wSource         As Workbook
Dim wDestination    As Workbook
Dim sheet           As Worksheet
Dim DateString      As String
Dim FolderName      As String
Dim i               As Integer
Dim rStart          As range
Dim rEnd            As range

With Application

    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
    
End With

Set wSource = ThisWorkbook

Sheets("passback").Activate

range("L1:L2").Replace What:="/", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Set rStart = range("L1")
Set rEnd = range("L2")
        
DateString = Format(Now, "mm_dd_yy")
FolderName = wSource.path & "\" & "Passback Templates " & range("L2").Value

On Error Resume Next
Kill FolderName & "\*.*"
RmDir FolderName
On Error GoTo 0

MkDir FolderName

Sheets(10).Activate

For i = 10 To Sheets.Count
    Sheets(i).Select (False)
Next

For Each sheet In ActiveWindow.SelectedSheets

    If sheet.Visible = -1 Then
        
        sheet.Copy
        
        Set wDestination = ActiveWorkbook
        
        With wDestination
        
            If Val(Application.Version) < 12 Then
            
                Filestr = ".xls": FileFormat = -4143
            
            Else
            
                If wSource.Name = .Name Then
                    GoTo GoToNextSheet
                Else
                    Select Case wSource.FileFormat
                        Case 51: Filestr = ".xlsx": FileFormat = 51
                        Case 52:
                            If .HasVBProject Then
                                Filestr = ".xlsm": FileFormat = 52
                            Else
                                Filestr = ".xlsx": FileFormat = 51
                            End If
                        Case 56: Filestr = ".xls": FileFormat = 56
                        Case Else: Filestr = ".xlsb": FileFormat = 50
                    End Select
                End If
            End If
        End With

Application.DisplayAlerts = False
        
        With wDestination
        
            Call Format_Templates
            
            .SaveAs FolderName _
                & "\" & wDestination.Sheets(1).range("C2") & "_" & rStart & " - " & rEnd & Filestr, _
                    FileFormat:=FileFormat
            .Close False
        End With
    End If
GoToNextSheet:
    Next sheet
    
Application.DisplayAlerts = True

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationManual
    End With
    
End Sub
