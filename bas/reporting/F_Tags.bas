Attribute VB_Name = "F_Tags"
Option Explicit

Sub F_Tag_URLs()

' Activate the F Tags sheet

Sheets("F_Tags").Activate

Range("E1", Range("E1").End(xlDown)).Select

' Some of the URLs in the F Tag sheet have extensions at the end, which sometimes
' causes the lookup to not bring in the correct URL for F Tags. The below searches
' for these extension strings and removes them

Selection.Replace What:=".html", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

Selection.Replace What:=".aspx", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

End Sub
