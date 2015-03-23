Attribute VB_Name = "Main"
Option Explicit

Sub DFA_Reporting()

Dim rDDR    As Range

Call Process_Raw_Reports
Call F_Tag_URLs

Call Python_Weekly_Reporting

Sheets("data").Activate
Columns("C:C").Select
Set rDDR = Selection.Find(What:="DDR", After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)

If rDDR Is Nothing Then

    Call Postprocess_Report

Else
    
    Call DDR_Top_15_Devices
    Call Postprocess_Report
    
End If

End Sub

Sub Generate_Passback_Templates()

Call PreRun

RunPython ("import passback_template_generation; passback_template_generation.passback_generation()")

Call SplitSheets

RunPython ("import email_passback_templates; email_passback_templates.email_passback_templates()")

End Sub

Sub Merge_Received_Templates()

Dim path    As Variant

path = InputBox("Enter path of received passback templates")

Sheets("Lookup").Activate

Range("AB1").Value = path

End Sub
