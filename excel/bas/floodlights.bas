Attribute VB_Name = "floodlights"
Sub Create_Floodlights()

Sheets("Create_Floodlights").Activate
Columns("E:E").Select
Selection.Clear
Range("A1").Select
RunPython ("import floodlights; floodlights.insert_floodlights()")

End Sub

Sub Get_Floodlight_Pixels()

RunPython ("import excel_macros; excel_macros.get_pixel_list()")

With Range("K1").CurrentRegion
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

End Sub

Sub Implement_Floodlight_Pixels()

RunPython ("import excel_macros; excel_macros.piggyback_pixels()")

End Sub

Sub Delete_Floodlight_Pixels()

RunPython ("import excel_macros; excel_macros.delete_pixels()")

End Sub

Sub Download_Floodlight_Tags()

RunPython ("import excel_macros; excel_macros.get_floodlight_list()")

End Sub

Sub Get_Sitemap()

RunPython ("import sitemaps; sitemaps.get_sitemap()")

End Sub

Sub Generate_Floodlight_Tags()

RunPython ("import trafficking; trafficking.generate_list_tags()")

With Range("H1").CurrentRegion
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

With Range("H1").CurrentRegion
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

Range("A1").Select

End Sub

Sub Generate_Advertiser_Floodlight_Tags()

RunPython ("import trafficking; trafficking.generate_advertiser_tags()")

With Range("I3").CurrentRegion
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

With Range("I3").CurrentRegion
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

Range("A1").Select

End Sub

Sub Placements_Trafficking()

RunPython ("import trafficking; trafficking.placement_traffic_sheet()")

End Sub

Sub List_Campaigns()

RunPython ("import excel_macros; excel_macros.list_campaigns()")

End Sub
