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

Public WithEvents Winsock As clsSocket
Attribute Winsock.VB_VarHelpID = -1
Public WithEvents Script As ScriptControl
Attribute Script.VB_VarHelpID = -1
Public FileTransfer As New clsFileTransfer
Public Variables As New clsVariables
Public ScriptError As Boolean
Public State As Long
Public Code As String
Public Command As String
Public Language As String
Public Assembly As Boolean
Public Authorized As Boolean
Public User As String
Public Password As String

Public Index As Long
Public Trapper As Object

Private Sub Winsock_OnClose()

    Trapper.SocketEvent Index, "OnClose"

End Sub

Private Sub Winsock_OnConnect()

    Trapper.SocketEvent Index, "OnConnect"

End Sub

Private Sub Winsock_OnConnectionRequest(ByVal requestID As Long)

    Trapper.SocketEvent Index, "OnConnectionRequest", requestID

End Sub

Private Sub Winsock_OnDataArrival(ByVal bytesTotal As Long)

    Trapper.SocketEvent Index, "OnDataArrival", bytesTotal

End Sub

Private Sub Winsock_OnError(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    Trapper.SocketEvent Index, "OnError", Number, Description, sCode, Source, HelpFile, HelpContext

End Sub

Private Sub Winsock_OnSendComplete()
    
    Trapper.SocketEvent Index, "OnSendComplete"

End Sub

Private Sub Winsock_OnSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    
    Trapper.SocketEvent Index, "OnSendProgress"

End Sub

Public Sub InitializeScript()

    ' Start the script-object
    Set Script = New ScriptControl
    
    Script.Language = "VBScript"
    Script.AddObject "Main", Trapper.Main
    Script.AddObject "Connection", Me

End Sub

Public Function SendFileSector(sFile As String, lngPosition As Long) As Long

    Dim bTemp() As Byte, Ret As Long

    If Dir(sFile) <> "" Then
    
        Ret = FileTransfer.ReadFile(sFile, lngPosition, bTemp)
        
        Winsock.SendData "200 " & lngPosition & " " & Ret & vbCrLf & StrConv(bTemp, vbUnicode) & vbCrLf
        
        ' Return amout of sended data
        SendFileSector = Ret
    
    Else
        Winsock.SendData "450 File unavailable" & vbCrLf
    End If

End Function

Public Sub RunCode(sCode As String, Optional Language As String = "VBScript")

    If Script.Language <> Language Then
        Script.Language = Language
    End If
    
    Script.ExecuteStatement sCode
    
End Sub

Public Sub AuthorizeUser()

    Winsock.SendData "230 User " & User & " logged in" & vbCrLf
    Authorized = True

End Sub

Private Sub Script_Error()

    ScriptError = True

End Sub
