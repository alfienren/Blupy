Attribute VB_Name = "External_Calls"
Option Explicit

Sub Python_Weekly_Reporting()

Sheets("Lookup").Activate
Range("AA1").Value = ActiveWorkbook.FullName

RunPython ("import main; main.weekly_reporting()")

End Sub

Sub Python_DDR_Top_Devices()

RunPython ("import ddr_weekly_reporting; ddr_weekly_reporting.ddr_top_15_devices()")

End Sub

Sub Python_Compress_Data()

RunPython ("import main; main.data_compression()")

End Sub

Sub Python_Split_Data()

RunPython ("import main; main.data_split()")

End Sub

Sub Python_Merge_Data()

RunPython ("import main; main.data_merge()")

End Sub

Sub Python_eBay_CostFeed()

RunPython ("import main; main.ebay_costfeed()")

End Sub
