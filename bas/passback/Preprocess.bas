Attribute VB_Name = "Preprocess"
Option Explicit

Sub PreRun()

Sheets("passback_placements").Activate

range("F1:F2").Replace What:=".", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
range("AA1").Value = ActiveWorkbook.FullName

End Sub
