Attribute VB_Name = "FileSelect"
Option Private Module
Option Explicit

Sub Select_Feed_File()

Set sFile = Application.FileDialog(msoFileDialogFilePicker)

With sFile

    .Title = "Select downloaded device feed text file"
    
    .AllowMultiSelect = False
    
    If .Show <> -1 Then
        Exit Sub
    End If
    
    FileSelected = .SelectedItems(1)
    
End With

Sheets("Action_Reference").Activate
Range("AE1").Value = FileSelected

End Sub
