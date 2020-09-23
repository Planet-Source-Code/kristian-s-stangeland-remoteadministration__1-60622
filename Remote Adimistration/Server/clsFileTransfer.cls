VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileTransfer"
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

Public FileName As String
Public FilePos As Long
Public FileByteSent As Long
Public FileByteToSend As Long
Public FileRemoveWhenSend As Boolean
Public FileOverflow As Long

Public Function GetDrives() As Variant

    On Error Resume Next
    Dim lngRet As Long
    Dim strDrives As String * 255
    Dim lngTmp As Long
    
    lngTmp = Len(strDrives)
    lngRet = GetLogicalDriveStrings(lngTmp, strDrives)
    GetDrives = Split(Left(strDrives, lngRet - 1), vbNullChar)
    
End Function

Public Function GetUNCPath(strDriveLetter As String) As String

    On Local Error GoTo fGetUNCPath_Err
    Dim Msg As String, lngReturn As Long
    Dim lpszLocalName As String
    Dim lpszRemoteName As String
    Dim cbRemoteName As Long
    
    lpszLocalName = strDriveLetter
    lpszRemoteName = String$(255, Chr$(32))
    cbRemoteName = Len(lpszRemoteName)
    lngReturn = WNetGetConnection(lpszLocalName, lpszRemoteName, cbRemoteName)
    
    Select Case lngReturn
        Case ERROR_BAD_DEVICE
            Msg = "Error: Bad Device"
            
        Case ERROR_CONNECTION_UNAVAIL
            Msg = "Error: Connection Un-Available"
            
        Case ERROR_EXTENDED_ERROR
            Msg = "Error: Extended Error"
            
        Case ERROR_MORE_DATA
               Msg = "Error: More Data"
               
        Case ERROR_NOT_SUPPORTED
               Msg = "Error: Feature not Supported"
               
        Case ERROR_NO_NET_OR_BAD_PATH
               Msg = "Error: No Network Available or Bad Path"
               
        Case ERROR_NO_NETWORK
               Msg = "Error: No Network Available"
               
        Case ERROR_NOT_CONNECTED
               Msg = "Error: Not Connected"
               
        Case NO_ERROR
               ' All is successful
               
    End Select
    
    If Len(Msg) = 0 Then
        Communication.Main.InvokeError Msg
    Else
        GetUNCPath = Left$(lpszRemoteName, cbRemoteName)
    End If
    
Exit Function
fGetUNCPath_Err:

    Communication.Main.InvokeError Err.Description
    
End Function

Public Function DriveType(strDriveName As String) As String

    Dim lngRet As Long
    Dim strDrive As String
    
    lngRet = GetDriveType(strDriveName)
    
    Select Case lngRet
        Case DRIVE_UNKNOWN 'The drive type cannot be determined.
            strDrive = "Unknown Drive Type"
            
        Case DRIVE_ABSENT 'The root directory does not exist.
            strDrive = "Drive does not exist"
            
        Case DRIVE_REMOVABLE 'The drive can be removed from the drive.
            strDrive = "Removable Media"
            
        Case DRIVE_FIXED 'The disk cannot be removed from the drive.
            strDrive = "Fixed Drive"
            
        Case DRIVE_REMOTE  'The drive is a remote (network) drive.
            strDrive = "Network Drive"
            
        Case DRIVE_CDROM 'The drive is a CD-ROM drive.
            strDrive = "CD Rom"
            
        Case DRIVE_RAMDISK 'The drive is a RAM disk.
            strDrive = "Ram Disk"
            
    End Select
    
    DriveType = strDrive
    
End Function

Public Function ReadFile(sFile As String, lngPosition As Long, bResult() As Byte) As Long

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

End Function

Public Function WriteFile(sFile As String, lngPosition As Long, bData() As Byte) As Long

    On Error GoTo Feil
    Dim Free As Long, lSize As Long
    
    Free = FreeFile
    
    ' Create the directory if it's not existing
    SHCreateDirectoryEx ByVal 0&, GetFilePath(sFile), ByVal 0&
    
    Open sFile For Binary Access Write Lock Write As #Free
        
        If lngPosition < 1 Then
            ' Default position is at the end of the file
            lngPosition = LOF(Free) + 1
        End If
        
        Seek #Free, lngPosition
        Put #Free, , bData
        
    Close #Free

    ' Return the size of the written data
    WriteFile = UBound(bData)

    Exit Function
Feil:
    
    Close #Free
    WriteFile = -1

End Function

Public Function ValidPath(sPath As String) As String

    ValidPath = sPath & IIf(Right(sPath, 1) = "\", "", "\")

End Function

Public Function GetFilePath(sFile As String) As String

    GetFilePath = Left(sFile, InStrRev(sFile, "\"))

End Function

Public Function FileSize(sFile As String) As Long

    FileSize = FileLen(sFile)

End Function

Public Function FileDate(sFile As String) As Date

    FileDate = FileDateTime(sFile)

End Function

Public Function FileAttributes(sFile As String) As Long

    FileAttributes = GetFileAttributes(sFile)

End Function

Public Sub DeleteFile(sFile As String)

    Kill sFile

End Sub

Public Function ListFiles(ByVal sDir As String, Optional Attributes As VbFileAttribute = vbNormal) As Variant

    Dim numFiles As Integer
    Dim aFiles() As Variant
    Dim sFile As String

    sFile = Dir$(sDir, Attributes)

    Do While Len(sFile) > 0

        If Not (sFile = "." Or sFile = "..") Then

            numFiles = numFiles + 1
            ReDim Preserve aFiles(1 To numFiles)
            
            ' Include different information about the file
            aFiles(numFiles) = sFile & "|" & FileLen(ValidPath(sDir) & sFile) & "|" & FileDateTime(ValidPath(sDir) & sFile)
            
        End If

        sFile = Dir$()
    Loop
    
    ListFiles = aFiles
    
End Function
