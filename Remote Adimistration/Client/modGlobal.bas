Attribute VB_Name = "modGlobal"
Option Explicit

Public Sub WriteFile(sFile As String, sData As String)

    Dim Free As Long
    
    Free = FreeFile
    
    Open sFile For Append As #Free
        Print #Free, sData
    Close #Free

End Sub
