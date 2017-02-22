Attribute VB_Name = "External_Calls"
Option Private Module
Option Explicit

Sub TMO_Weekly_Reporting()

RunPython ("import tmo; tmo.generate_weekly_reporting()")

End Sub

Sub Python_eBay_CostFeed()

RunPython ("import tmo; tmo.cost_feed()")

End Sub

Sub Python_FlatRates()

RunPython ("import tmo; tmo.output_flat_rate_report()")

End Sub

Sub Python_Pacing_Report()

RunPython ("import tmo; tmo.pacing_report()")

End Sub

Sub Metro_Weekly_Reporting()

RunPython ("import main; main.generate_metro_reporting()")

End Sub
