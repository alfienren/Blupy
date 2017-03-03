
Sub TMO_Weekly_Reporting()

Dim rDDR    As Range

Call TMO_Raw_Reports

Call Select_Feed_File

RunPython ("import excel_macros; excel_macros.tmobile_weekly_reporting()")

Call TMO_Postprocess_Report

End Sub



Sub Metro_Weekly_Reporting()

RunPython ("import excel_macros; excel_macros.metro_weekly_reporting()")

End Sub



Sub WFM_Weekly_Reporting()

Call WFM_Raw_Reports
Call WFM_Postprocess_Report
RunPython ("import excel_macros; excel_macros.wfm_weekly_reporting()")

End Sub