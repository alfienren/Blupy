Attribute VB_Name = "Main"
Option Explicit

Sub DFA_Reporting()

Call Process_Raw_Reports
Call F_Tag_URLs

Call Postprocess_Report

End Sub