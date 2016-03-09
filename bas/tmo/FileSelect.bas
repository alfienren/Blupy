Attribute VB_Name = "FileSelect"
Option Private Module

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

Sub Select_Trafficking_Campaign_Master_File()

Set sFile = Application.FileDialog(msoFileDialogFilePicker)

With sFile

    .Title = "Select trafficking Campaign Master File"
    
    .AllowMultiSelect = False
    
    If .Show <> -1 Then
        Exit Sub
    End If
    
    FileSelected = .SelectedItems(1)
    
End With

Sheets("Action_Reference").Activate
Range("AE1").Value = FileSelected

End Sub

Sub Select_Campaign_Trafficking_Reports_Folder()

Set sFile = Application.FileDialog(msoFileDialogFolderPicker)

With sFile

    .Title = "Select folder containing campaign traffic reports to merge into master"
    
    .AllowMultiSelect = False
    
    If .Show <> -1 Then
        Exit Sub
    End If
    
    FileSelected = .SelectedItems(1)
    
End With

Sheets("Action_Reference").Activate
Range("AE1").Value = FileSelected

End Sub
