Attribute VB_Name = "Preprocess"
Option Explicit

Sub PreRun()

Sheets("passback").Activate

range("L1:L2").Replace What:=".", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
range("AA1").Value = ActiveWorkbook.FullName

End Sub
