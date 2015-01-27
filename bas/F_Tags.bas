Attribute VB_Name = "F_Tags"
Option Explicit

Sub F_Tag_URLs()

Sheets("F_Tags").Activate

Range("E1", Range("E1").End(xlDown)).Select

Selection.Replace What:=".html", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

End Sub
