Attribute VB_Name = "External_Calls"
Option Explicit

Sub Python_Weekly_Reporting()

Sheets("Action_Reference").Activate
Range("AG1").Value = ActiveWorkbook.FullName

RunPython ("import main; main.weekly_reporting()")

End Sub

'Sub Python_DDR_Top_Devices()
'
'RunPython ("import main; main.dr_device_report()")
'
'End Sub

Sub Python_Compress_Data()

RunPython ("import main; main.data_compression()")

End Sub

Sub Python_Split_Data()

RunPython ("import main; main.data_split()")

End Sub

Sub Python_Merge_Data()

RunPython ("import main; main.data_merge()")

End Sub

Sub Python_TMO_CostFeed()

RunPython ("import main; main.tmo_costfeed()")

End Sub
