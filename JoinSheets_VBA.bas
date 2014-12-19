Attribute VB_Name = "JoinSheets_VBA"
Option Explicit

Sub JoinData()





'Dim oCONN           As ADODB.Connection
'Dim oRS             As ADODB.Recordset
'Dim wsOutput        As Worksheet
'Dim strSQL          As String
'Dim strConn         As String
'Dim i               As Long
'Dim hdrName         As Variant
'
'Set wsOutput = ActiveWorkbook.Worksheets("working")
'
'wsOutput.Cells.ClearContents
'
'Set oCONN = New ADODB.Connection
'Set oRS = New ADODB.Recordset
'
'strConn = "Provider=Microsoft.ACE.OLEDB.15.0;" & _
'            "Data Source=" & ActiveWorkbook.FullName & ";" & _
'            "Extended Properties=""Excel 12.0;HDR=Yes;"";"
'
'strSQL = "SELECT[SA_Temp$].* " & _
'        "FROM [SA_TEMP$] INNER JOIN [CFV_Temp$] " & _
'        "ON [SA_Temp$].UniqueID = [CFV_Temp$].UniqueID"
'
'oCONN.Open strConn
'oRS.Open strSQL, oCONN
'
'i = 1
'
'For Each hdrName In oRS.Fields
'
'    wsOutput.Cells(1, i).Value = hdrName.Name
'    i = i + 1
'
'Next hdrName
'
'wsOutput.Cells(2, 1).CopyFromRecordset oRS
'
'With wsOutput
'
'    .Activate
'    .ListObjects.Add(xlSrcRange, ActiveSheet.UsedRange, , xlYes).Name = "Tbl_Output"
'    .Cells.EntireColumn.AutoFit
'
'End With
'
'oCONN.Close
'Set oCONN = Nothing
'Set oRS = Nothing

End Sub
