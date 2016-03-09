Attribute VB_Name = "External_Calls"
Option Private Module
Option Explicit

Sub Python_Weekly_Reporting()

Sheets("Action_Reference").Activate
Range("AG1").Value = ActiveWorkbook.FullName

RunPython ("import main; main.generate_weekly_reporting()")

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

RunPython ("import main; main.tmo_costfeed()")

End Sub

Sub Python_FlatRates()

RunPython ("import main; main.flat_rate_costs()")

End Sub
