Attribute VB_Name = "Main"
Option Explicit

Sub DFA_Reporting()

Call Process_Raw_Reports ' Create necessary temporary tabs and enter in SA and CFV data
Call F_Tag_URLs ' Make sure URLs are friendly for lookups

Call Python_Calls ' Run the Python module

Call Postprocess_Report ' Clean up the sheet by deleting the temporary tabs and any extraneous references

End Sub
