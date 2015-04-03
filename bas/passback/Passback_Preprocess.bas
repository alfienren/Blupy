Attribute VB_Name = "Passback_Preprocess"
Option Explicit

Sub PreRun()

Sheets("passback").Activate

Range("L1:L2").Replace What:=".", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
Range("AA1").Value = ActiveWorkbook.FullName

End Sub
