VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommunication"
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

' Objects
Private Connections() As clsConnection
Public WithEvents PrimaryConnection As clsSocket
Attribute PrimaryConnection.VB_VarHelpID = -1
Public CurrentConnection As clsConnection
Public Main As New clsMain

Public Property Get Clients(ByVal Index As Long) As clsConnection

    If IsClient(Index) Then
        Set Clients = Connections(Index)
    End If

End Property

Public Property Let Clients(ByVal Index As Long, ByVal vNewValue As clsConnection)

    If IsClient(Index) Then
        Connections(Index) = vNewValue
    End If

End Property

Public Function IsClient(ByVal Index As Long) As Boolean

    IsClient = CBool(Index >= LBound(Connections) And Index <= UBound(Connections))

End Function

Public Sub SendScreenshot(Connection As clsConnection, Extension As String, Optional Width As Long, Optional Height As Long, Optional srcWidth As Long, Optional srcHeight As Long)

    Dim sFile As String

    sFile = Main.GDIPlus.MakeScreenshot(Extension, Width, Height, srcWidth, srcHeight)

    If Dir(sFile) = "" Then
        Connection.Winsock.SendData "452 Error making screenshot"
    Else
        SendFile Connection, sFile, 1
    End If

End Sub

Public Sub SendFile(Connection As clsConnection, sFile As String, lngPosition As Long)

    Dim Ret As Long

    Connection.State = 3
    
    Connection.FileTransfer.FileName = sFile
    Connection.FileTransfer.FileByteSent = 0
    Connection.FileTransfer.FileByteToSend = FileLen(sFile)
    Connection.FileTransfer.FilePos = IIf(lngPosition < 1, 1, lngPosition)
    
    ' Send the first sector
    Ret = Connection.SendFileSector(sFile, Connection.FileTransfer.FilePos)
    
    If Ret < 0 Then
        Connection.State = 0
        Connection.Winsock.SendData "451 Error sending file" & vbCrLf
    Else
        Connection.FileTransfer.FilePos = lngPosition + Ret
    End If

End Sub

Private Function GetElement(aArray As Variant, Index As Long) As Variant

    On Error Resume Next
    GetElement = aArray(Index)

End Function

Public Sub ExecuteData(Connection As clsConnection, sData As String)

    On Error Resume Next
    Dim aBuffer, aTemp, Tell As Long, bTemp() As Byte, Ret As Long, sPass As String
    Dim sFile As String, ProcessID As Long, hProcess As Long, sTemp As String
    
    If sData = "" Then
        ' We have nothing to execute
        Exit Sub
    End If
    
    ' Get the information
    aBuffer = Split(Connection.Command & sData, vbCrLf)
    
    For Tell = LBound(aBuffer) To UBound(aBuffer)
    
        Select Case Connection.State
        Case 0 ' Listen for commands
        
            If UBound(aBuffer) > 0 Then
            
                ' Split the command up in pieces
                aTemp = Main.SplitX(CStr(aBuffer(Tell)), " ")
                
                If UBound(aTemp) >= 0 Then
                
                    If Connection.Authorized = False Then
                        
                        Select Case LCase(aTemp(0))
                        Case "user"
                        
                            Connection.User = Main.RemoteQuote(aTemp(1))
                            sPass = Main.UserPassword(Connection.User)
                        
                            If sPass <> "" Then
                            
                                ' Send request for password
                                Connection.Winsock.SendData "331 Password requried for " & Connection.User & vbCrLf
                            
                            Else
                                
                                 ' User logged in
                                Connection.AuthorizeUser
                            
                            End If
                        
                        Case "pass"
                        
                            Connection.Password = Main.RemoteQuote(aTemp(1))
                            
                            If Main.UserPassword(Connection.User) = Connection.Password And Connection.Password <> Chr(34) & "ERROR" & Chr(34) Then
                            
                                ' User logged in
                                Connection.AuthorizeUser
                            
                            Else
                            
                                ' The login wasn't successful
                                Connection.Winsock.SendData "530 Login incorrect" & vbCrLf
                            
                            End If
                        
                        Case "quit", "disconnect", "close", "exit"
                        
                            ' Close this connection
                            CloseConnection Connection
                        
                        Case Else
                            
                            ' Tell the user to log on
                            Connection.Winsock.SendData "530 Please log in with USER or PASS." & vbCrLf
                        
                        End Select
                
                    Else
                
                        ' Execute the command
                        Select Case LCase(aTemp(0))
                        Case "user", "pass"
                        
                            ' Simply tell the user that it's authorized
                            Connection.AuthorizeUser
                        
                        Case "logout"
                        
                            ' We're loging out
                            Connection.Authorized = False
                            
                            ' Repond to this action
                            Connection.Winsock.SendData "200 User logged out" & vbCrLf
                        
                        Case "runcode"
                        
                            ' Clear the last VB-code
                            Connection.Code = ""
                        
                            ' Wait for the code to run
                            Connection.State = 1
                            
                            ' This is not assembly but script
                            Connection.Assembly = False
                            
                            ' Get the language
                            If UBound(aTemp) > 0 Then
                                Connection.Language = aTemp(1)
                            Else
                                Connection.Language = "VBScript" ' Default
                            End If
                            
                            ' Repond to the client
                            Connection.Winsock.SendData "150 Waiting for data" & vbCrLf
                        
                        Case "runexe"
                        
                            ' Get the file
                            If UBound(aTemp) > 0 Then
                            
                                sFile = Main.RemoteQuote(aTemp(1))
                                sTemp = Main.ReturnOutput(sFile)
                                
                                Connection.Winsock.SendData "200 1 " & Len(sTemp) & vbCrLf & sTemp & vbCrLf
                            
                            Else
                                Connection.Winsock.SendData "501 Wrong number of paramenters." & vbCrLf
                            End If
                        
                        Case "runasm"
                        
                            ' Initialize variables
                            Connection.Code = ""
                            Connection.State = 1
                            Connection.Assembly = True
                    
                            ' Repond to the client
                            Connection.Winsock.SendData "150 Waiting for data" & vbCrLf
                        
                        Case "shell"
                        
                            If UBound(aTemp) = 1 Then
    
                                ProcessID = Shell(Main.RemoteQuote(aTemp(1)))
                                
                                If aTemp(2) = 1 Then
                                    hProcess = OpenProcess(SYNCHRONIZE, 0, ProcessID)
                                    If hProcess <> 0 Then
                                        WaitForSingleObject hProcess, INFINITE
                                        CloseHandle hProcess
                                    End If
                                End If
                                
                                Connection.Winsock.SendData "200 Command executed." & vbCrLf
                            Else
                                Connection.Winsock.SendData "501 Wrong number of paramenters." & vbCrLf
                            End If
                        
                        Case "screenshot"
                        
                            ' Send the screenshot
                            If UBound(aTemp) >= 1 Then
                                Connection.FileTransfer.FileRemoveWhenSend = True
                                SendScreenshot Connection, CStr(aTemp(1)), GetElement(aTemp, 2), GetElement(aTemp, 3), GetElement(aTemp, 4), GetElement(aTemp, 5)
                            Else
                                Connection.Winsock.SendData "501 Wrong number of paramenters." & vbCrLf
                            End If
                        
                        Case "quit", "disconnect", "close", "exit"
                        
                            ' Close this connection
                            CloseConnection Connection
                        
                        Case "exitapp"
                        
                            ' Exit the application
                            ExitApp = True
                        
                        Case "listdrives"
    
                            sTemp = Join(Connection.FileTransfer.GetDrives, vbCrLf)
                
                            ' Send the list
                            Connection.Winsock.SendData "200 Sending drive-list, " & Len(sTemp) & " bytes" & vbCrLf & sTemp & vbCrLf
                        
                        Case "mdtm"
                        
                            ' Get the file
                            sFile = Main.RemoteQuote(aTemp(1))
                        
                            If Dir(sFile) <> "" Then
                                ' Send the file date time
                                Connection.Winsock.SendData "213 " & FileDateTime(sFile) & vbCrLf
                            Else
                                Connection.Winsock.SendData "450 File unavailable" & vbCrLf
                            End If
                        
                        Case "list"
                        
                            ' Get the path
                            sFile = Connection.FileTransfer.ValidPath(Main.RemoteQuote(aTemp(1)))
                            
                            If UBound(aTemp) > 1 Then
                                Ret = Val(aTemp(2))
                            Else
                                Ret = vbNormal
                            End If
                            
                            ' Get the file-list
                            sTemp = Join(Connection.FileTransfer.ListFiles(sFile, Ret), vbCrLf)
                            
                            ' Send the list
                            Connection.Winsock.SendData "200 Sending output" & vbCrLf & sTemp & vbCrLf
                        
                        Case "delete"
                        
                            Kill Main.RemoteQuote(aTemp(1))
                        
                            ' Send response
                            Connection.Winsock.SendData "200 File deleted" & vbCrLf
                        
                        Case "rename"
                        
                            If UBound(aTemp) = 2 Then
                            
                                ' Change name
                                Name Main.RemoteQuote(aTemp(1)) As Main.RemoteQuote(aTemp(2))
                            
                                ' Send response
                                Connection.Winsock.SendData "250 File renamed" & vbCrLf
                            
                            Else
                                Connection.Winsock.SendData "501 Wrong number of paramenters." & vbCrLf
                            End If
                        
                        Case "rmdir"
                        
                            ' Remove the directory
                            RmDir Main.RemoteQuote(aTemp(1))
                        
                            ' Send response
                            Connection.Winsock.SendData "251 Directory removed" & vbCrLf
                        
                        Case "mdir"
                        
                            ' Create a directory
                            SHCreateDirectoryEx ByVal 0&, Main.RemoteQuote(aTemp(1)), ByVal 0&
                        
                            ' Send response
                            Connection.Winsock.SendData "250 Directory created" & vbCrLf
                        
                        Case "get", "recv"
                        
                            If UBound(aTemp) = 2 Then
                                Connection.FileTransfer.FileRemoveWhenSend = False
                                SendFile Connection, Main.RemoteQuote(aTemp(1)), Val(aTemp(2))
                            Else
                                Connection.Winsock.SendData "501 Wrong number of paramenters." & vbCrLf
                            End If
                        
                        Case "put"
                        
                            If UBound(aTemp) = 3 Then
                                Connection.FileTransfer.FileName = Main.RemoteQuote(aTemp(1))
                                Connection.FileTransfer.FilePos = Val(aTemp(2))
                                Connection.FileTransfer.FileByteToSend = Val(aTemp(3))
                                Connection.FileTransfer.FileByteSent = 0
                                Connection.State = 2
                            Else
                                Connection.Winsock.SendData "501 Wrong number of paramenters." & vbCrLf
                            End If
                        
                        Case "noop"
                        
                            Connection.Winsock.SendData "200 NOOP command successful." & vbCrLf
                        
                        Case Else
                        
                            Connection.Winsock.SendData "500 " & Chr(34) & aTemp(0) & Chr(34) & " command not recognized" & vbCrLf
                        
                        End Select
                    
                    End If
                    
                    ' Reset the buffer
                    Connection.Command = ""
                
                End If
            
            Else
            
                Connection.Command = Connection.Command & sData
            
            End If
            
        Case 1 ' Receive VB-code
        
            ' If the line is empty, we have ended the transmission
            If aBuffer(Tell) = "" And Connection.Code <> "" Then
            
                ' This is the current connection
                Set CurrentConnection = Connection
            
                ' No script error
                Connection.ScriptError = False
                
                If Connection.Assembly Then
                
                    HexToArray Connection.Code, bTemp
                    CallWindowProc VarPtr(bTemp(0)), 0, 0, 0, 0
                
                Else ' If this is script
            
                    ' Simply run the code
                    Connection.RunCode Connection.Code, Connection.Language
                    
                End If
                
                ' Inform the client
                If Connection.ScriptError = True Then
                    Connection.Winsock.SendData "500 Script error " & Err & vbCrLf
                Else
                    Connection.Winsock.SendData "200 Code executed" & vbCrLf
                End If
                
                ' Reset the state
                Connection.State = 0
                
                ' Exit the prosedure
                Exit Sub
            End If
        
            ' Add the line, together with the line break
            Connection.Code = Connection.Code & aBuffer(Tell) & IIf(Tell < UBound(aBuffer), vbCrLf, "")

        Case 2 ' Write to file

            ReadWritePackage Connection, sData

        Case 3 ' Send file
        
            ' Do nothing

        End Select

    Next

End Sub

Public Sub ReadWritePackage(Connection As clsConnection, sData As String)

    Dim sCode As String, aCode As Variant, Tell As Long, Bytes As Long, DataStart As Long, bTemp() As Byte

    ' Get a byte-array
    bTemp = sData
    
    ' We always starts reading from the beginning of the package
    Tell = 1

    ' If the data from the last segment wasn't long enough, assume this is the rest of it
    If Connection.FileTransfer.FileOverflow > 0 Then
    
        ' Rezie the byte-array, so it only write the data that it's supposed to write
        ReDim Preserve bTemp(Connection.FileTransfer.FileOverflow)
    
        ' Write the data
        Connection.FileTransfer.WriteFile Connection.FileTransfer.FileName, 0, bTemp
        
        ' Increse and reset variables
        Tell = Connection.FileTransfer.FileOverflow + 3
        Connection.FileTransfer.FileOverflow = 0
        
    End If

    ' Start the reading-loop
    Do Until Tell > Len(sData)
    
        ' Firstly we need to find out were the information header ends
        DataStart = InStr(Tell, sData, vbCrLf)
        
        ' If it's there, proceed
        If DataStart > 0 Then
        
            ' Split the information header into a array
            aCode = Split(Mid(sData, Tell, DataStart - Tell), " ", 3)
            
            ' Save the amout of bytes for later use
            Bytes = Val(aCode(2))
        
            ' Do the appropriate action for each trasnfer-state
            If Val(aCode(0)) = 200 Then
            
                ' Check if the package isn't long enough
                If Len(sData) - Bytes - DataStart - 2 < 1 Then
                    Connection.FileTransfer.FileOverflow = Bytes - Len(sData) + DataStart + 1
                    Bytes = Bytes - Connection.FileTransfer.FileOverflow
                End If
            
                ' Retrive a byte array for the data to be written
                bTemp = Mid(sData, DataStart + 2, Bytes)
            
                ' Write to the file
                Connection.FileTransfer.WriteFile Connection.FileTransfer.FileName, Val(aCode(1)), bTemp
                
                ' Increse the variable holding the position of the package
                Tell = DataStart + Bytes + 4
            
            ElseIf Val(aCode(0)) = 255 Then
                ' We're finished
                Connection.State = 0
                Exit Do
            Else
                ' This is a critical error
                Connection.Winsock.SendData "500 Transfer-state couldn't be read."
            
                ' Exit the state of reciving data
                Connection.State = 0
                Exit Do
            End If
        Else
            Exit Do
        End If
        
    Loop

End Sub

Public Sub StartServer()

    Set PrimaryConnection = New clsSocket

    PrimaryConnection.LocalPort = 8000
    PrimaryConnection.Listen

End Sub

Public Sub StopServer()

    ' Close the primary socket
    If Not (PrimaryConnection Is Nothing) Then
        PrimaryConnection.CloseSocket
        Set PrimaryConnection = Nothing
    End If
    
End Sub

Private Sub Class_Initialize()

    ' Give us one free socket
    ReDim Connections(0)

End Sub

Private Sub PrimaryConnection_OnConnectionRequest(ByVal requestID As Long)

    Dim Tell As Long
    
    For Tell = LBound(Connections) To UBound(Connections)
        If Not (Connections(Tell) Is Nothing) Then
            If Connections(Tell).Winsock Is Nothing Then
                Exit For
            Else
                If (Connections(Tell).Winsock.State = sckClosed) Or (Connections(Tell).Winsock.State = sckError) Then
                    Exit For
                End If
            End If
        End If
    Next
    
    If Tell > UBound(Connections) Then
        ReDim Preserve Connections(Tell)
    End If
    
    ' Initialize class-modules
    Set Connections(Tell) = New clsConnection
    Set Connections(Tell).Winsock = New clsSocket
    Set Connections(Tell).Trapper = Me
    
    ' See if user-authorization is neccessary
    Connections(Tell).Authorized = Not Main.UsersLogon
    
    ' Give the class knowledge about which index it is
    Connections(Tell).Index = Tell
    
    ' Initialize the script engine
    Connections(Tell).InitializeScript
    
    ' We're always going to connect
    Connections(Tell).Winsock.CloseSocket
    Connections(Tell).Winsock.Accept requestID
    
    ' Send the start-up message
    Connections(Tell).Winsock.SendData Connections(Tell).Winsock.LocalHostName & " RAC-server (Version " & App.Major & "." & App.Minor & "." & App.Revision & " " & Now & ") ready." & vbCrLf
    
End Sub

Public Sub CloseConnection(Connection As clsConnection)

    ' Close the socket
    Connection.Winsock.CloseSocket
    
    ' Clear all assosiated timers
    TimerObject.RemoveByConnection Connection.Index

    ' Free all resources
    Set Connection = Nothing

End Sub

Public Sub SocketEvent(Index As Long, EventType As String, ParamArray Paramenters() As Variant)

    On Error Resume Next
    Dim sBuffer As String, Ret As Long

    Select Case EventType
    Case "OnClose"
    
        ' Close this connection
        CloseConnection Connections(Index)
    
    Case "OnSendComplete"
    
        ' If we're about to send a file
        If Connections(Index).State = 3 Then
        
            If Connections(Index).FileTransfer.FilePos < Connections(Index).FileTransfer.FileByteToSend Then
                
                ' Send the next sector
                Ret = Connections(Index).SendFileSector(Connections(Index).FileTransfer.FileName, Connections(Index).FileTransfer.FilePos)
                
                If Ret < 0 Then
                    Connections(Index).State = 0
                    Connections(Index).Winsock.SendData "451 Error sending file" & vbCrLf
                Else
                    Connections(Index).FileTransfer.FilePos = Connections(Index).FileTransfer.FilePos + Ret
                End If
            
            Else
            
                If Connections(Index).FileTransfer.FileRemoveWhenSend Then
                    Kill Connections(Index).FileTransfer.FileName
                End If
            
                Connections(Index).State = 0
                Connections(Index).Winsock.SendData "255 File send" & vbCrLf
            End If
        
        End If
    
    Case "OnDataArrival"
    
        Connections(Index).Winsock.GetData sBuffer
        ExecuteData Connections(Index), sBuffer
    
    End Select
    
End Sub
