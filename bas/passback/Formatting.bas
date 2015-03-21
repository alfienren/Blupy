Attribute VB_Name = "Formatting"
Option Explicit

Sub Format_Templates()

Dim rtemplate   As range

range("A1", range("A1").End(xlToRight)).EntireColumn.AutoFit

Set rtemplate = range("A1").CurrentRegion

rtemplate.Borders(xlDiagonalDown).LineStyle = xlNone
rtemplate.Borders(xlDiagonalUp).LineStyle = xlNone

With rtemplate.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With rtemplate.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With rtemplate.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With rtemplate.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With rtemplate.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
    
With rtemplate.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
    
range("A1", range("A1").End(xlToRight)).Font.Bold = True

With range("A1", range("A1").End(xlToRight)).Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = 0.599993896298105
    .PatternTintAndShade = 0
End With

End Sub
