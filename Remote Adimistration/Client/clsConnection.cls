VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' The winsock-control
Public WithEvents Winsock As clsSocket
Attribute Winsock.VB_VarHelpID = -1
Public ConnectError As Boolean
Public Connected As Boolean

' Events this class has
Event DataSending(sData As String)
Event DataArrival(sData As String)
Event DataSent()
Event ConnectionClosed()

Public Property Get Name() As String

    Name = "Connection"
    
End Property

Private Sub Class_Initialize()

    ' Create the new socket
    Set Winsock = New clsSocket

End Sub

Public Sub Connect(sRemoteHost As String, Optional Port As Long = 8000)

    Winsock.Connect sRemoteHost, Port

End Sub

Public Sub Disconnect()

    Winsock.CloseSocket

End Sub

Public Function ReadFile(sFile As String, lngPosition As Long, bResult() As Byte) As Long

    On Error GoTo Feil
    Dim Free As Long, lSize As Long
    
    Free = FreeFile
    
    Open sFile For Binary Access Read Lock Write As #Free
        
        Seek #Free, lngPosition
        lSize = LOF(Free) - lngPosition + 1
        
        If lSize > ChunckSize Then
            lSize = ChunckSize
        End If
        
        ReDim bResult(1 To lSize)
        Get #Free, , bResult
        
    Close #Free

    ' Return the size of the buffer
    ReadFile = lSize
    
    Exit Function
Feil:

    ' Notify about the error
    MsgBox "Error " & Err & " in ReadFile: " & Error$, vbCritical, "Error"

    ' No bytes were read
    ReadFile = -1

End Function

Public Sub SendData(varData)

    ' Inform about the sendning
    RaiseEvent DataSending(CStr(varData))

    ' Send the data
    Winsock.SendData varData

End Sub

Public Function SendFileSector(sFile As String, lngPosition As Long) As Long

    Dim bTemp() As Byte, Ret As Long

    If Dir(sFile) <> "" Then
    
        Ret = ReadFile(sFile, lngPosition, bTemp)
        
        If Ret >= 0 Then
            SendData "200 " & lngPosition & " " & Ret & vbCrLf & StrConv(bTemp, vbUnicode) & vbCrLf
        Else
            SendData "500 Error reading file"
        End If
        
        ' Return amout of sended data
        SendFileSector = Ret

    Else
        Winsock.SendData "450 File unavailable" & vbCrLf
    End If

End Function

Private Sub Winsock_OnClose()

    RaiseEvent ConnectionClosed

End Sub

Private Sub Winsock_OnConnect()

    Connected = True

End Sub

Private Sub Winsock_OnDataArrival(ByVal bytesTotal As Long)

    Dim sData As String
    
    ' Get the data
    Winsock.GetData sData
    
    ' Inform about the arrival
    RaiseEvent DataArrival(sData)

End Sub

Private Sub Winsock_OnError(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    MsgBox "Error " & Number & ": " & Description, vbCritical, "Error"

End Sub

Private Sub Winsock_OnSendComplete()

    ' Raise the event
    RaiseEvent DataSent

End Sub
