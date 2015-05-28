Attribute VB_Name = "ModuleCopy"
Option Explicit

Sub CopyModule(wSource As Workbook, _
                sModuleName As String, _
                wDest As Workbook)
                
Dim sFolder             As String
Dim sTemp               As String

sFolder = wSource.path

If Len(sFolder) = 0 Then sFolder = CurDir

sFolder = sFolder & "\"
sTemp = sFolder & "~tmpexport.bas"

On Error Resume Next

wSource.VBProject.VBComponents(sModuleName).Export sTemp
wDest.VBProject.VBComponents.Import sTemp

Kill sTemp

On Error GoTo 0

End Sub
