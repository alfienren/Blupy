Attribute VB_Name = "External_Calls"
Option Explicit

Sub Python_Calls()

RunPython ("import weekly_reporting; weekly_reporting.dfa_reporting()")

End Sub
