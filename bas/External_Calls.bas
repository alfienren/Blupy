Attribute VB_Name = "External_Calls"
Option Explicit

Sub Python_Weekly_Reporting()

Sheets("Lookup").Activate
Range("AA1").Value = ActiveWorkbook.FullName

RunPython ("import weekly_reporting; weekly_reporting.weekly_reporting()")

End Sub

Sub Python_DDR_Top_Devices()

RunPython ("import ddr_weekly_reporting; ddr_weekly_reporting.ddr_top_15_devices()")

End Sub

Sub Python_Compress_Data()

RunPython ("import tools; tools.data_compression()")

End Sub
