VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
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

Public Registry As Object
Public Message As New clsMessage
Public GDIPlus As New clsGdiPlus
Public Variables As New clsVariables
Public API As New clsFuncCall
Public IgnoreError As Boolean

Public Property Get Winsock() As clsSocket

    Set Winsock = New clsSocket

End Property

Public Property Get ScreenWindow() As Form

    Set ScreenWindow = frmScreen

End Property

Public Property Get Com() As clsCommunication

    Set Com = Communication

End Property

Public Property Get Timer() As clsTimer

    Set Timer = TimerObject

End Property

Public Property Get TickCount() As Long

    TickCount = GetTickCount

End Property

Public Property Get KeyState(VirtualKey As Long)

    KeyState = GetKeyState(VirtualKey)

End Property

Public Property Let LockInput(vNewValue As Boolean)

    BlockInput IIf(vNewValue, 1, 0)

End Property

Public Property Let CursorVisible(vNewValue As Boolean)

    ShowCursor vNewValue

End Property

Public Property Get CursorX() As Long

    Dim Pos As POINTAPI

    GetCursorPos Pos
    CursorX = Pos.X

End Property

Public Property Let CursorX(ByVal vNewPos As Long)

    SetCursorPos vNewPos, CursorY

End Property

Public Property Get CursorY() As Long

    Dim Pos As POINTAPI

    GetCursorPos Pos
    CursorY = Pos.Y

End Property

Public Property Let CursorY(ByVal vNewPos As Long)

    SetCursorPos CursorX, vNewPos

End Property

Public Property Get TempFile(Optional sPrefix As String = "RAC") As String

    Dim sFile As String, sPath As String

    sPath = String(100, 0)
    sFile = String(260, 0)

    GetTempPath 100, sPath
    GetTempFileName sPath, sPrefix, 0, sFile

    TempFile = Left(sFile, InStr(1, sFile, Chr(0)) - 1)

End Property

Public Sub DisableTaskman(Optional Value As Long = 1)

    Registry.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", Value

End Sub

Public Property Get StartWithWindows() As Boolean

    On Error Resume Next
    Dim sTemp As String
    
    ' Read the key
    sTemp = Registry.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName)

    StartWithWindows = (Len(sTemp) <> 0)

End Property

Public Property Let StartWithWindows(ByVal vNewValue As Boolean)

    Dim sKey As String
    
    ' The key we're going to write to or delete
    sKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName

    If vNewValue = True Then
        Registry.RegWrite sKey, App.Path & "\" & App.EXEName & ".exe"
    Else
        Registry.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName
    End If

End Property

Public Function EnumKeys(Key As Long, sSubKey As String, Optional Value As Boolean) As Variant

    Dim hKey As Long, Cnt As Long, sName As String, sData As String, Ret As Long, FuncRet As Long, aKeys()
    
    Ret = 255
    
    If RegOpenKey(Key, sSubKey, hKey) = 0 Then

        sName = Space(255)
        Ret = 255
        
        While FuncRet <> 259
            
            If Value Then
                FuncRet = RegEnumValue(hKey, Cnt, sName, Ret, ByVal 0&, 0&, ByVal 0&, ByVal 0&)
            Else
                FuncRet = RegEnumKeyEx(hKey, Cnt, sName, Ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&)
            End If
            
            If FuncRet <> 259 Then
            
                ' Save the enumerated key
                ReDim Preserve aKeys(Cnt)
                aKeys(Cnt) = Left$(sName, Ret)
                
                ' Prepare for the next key
                Cnt = Cnt + 1
                
                sName = Space(255)
                Ret = 255
                
            End If
        Wend
        
        ' Close the registry key
        RegCloseKey hKey
    Else
        'InvokeError "Error while calling RegOpenKey"
    End If

    ' Return the keys
    EnumKeys = aKeys

End Function

Public Sub DeleteUser(ByVal sUserName As String)
    
    On Error Resume Next
    DeleteSetting "RAC", "Users", sUserName

End Sub

Public Property Get UserPassword(ByVal sUserName As String) As String

    UserPassword = GetSetting("RAC", "Users", sUserName, Chr(34) & "ERROR" & Chr(34))

End Property

Public Property Let UserPassword(ByVal sUserName As String, ByVal sNewValue As String)

    SaveSetting "RAC", "Users", sUserName, sNewValue

End Property

Public Property Get UsersLogon() As Boolean

    UsersLogon = GetSetting("RAC", "General", "UsersLogon", False)

End Property

Public Property Let UsersLogon(ByVal vNewValue As Boolean)

    SaveSetting "RAC", "General", "UsersLogon", vNewValue

End Property

Public Function IsWinNT() As Boolean

    Dim myOS As OSVERSIONINFO
    
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
    
End Function

Public Function WindowsVersion() As String
    
    Dim Ver As Long, WinVer As Long
    
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    WindowsVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
    
End Function

Public Sub ExitWindows(Optional flags As Long, Optional Force As Boolean)

    flags = EWX_SHUTDOWN
    
    If Force = True Then
        flags = flags + EWX_FORCE
    End If
    
    If IsWinNT Then
        EnableShutDown
    End If
    
    ExitWindowsEx flags, 0

End Sub

Public Sub SendKey(bKey As Byte, Action As Long)

    Dim GInput As GENERALINPUT
    Dim KInput As KEYBDINPUT
    
    KInput.wVk = bKey
    KInput.dwFlags = Action
    GInput.dwType = INPUT_KEYBOARD
    CopyMemory GInput.xi(0), KInput, Len(KInput)

    'send the input now
    SendInput 1, GInput, Len(GInput)
    
End Sub

Public Sub InvokeError(sText As String)

    If IgnoreError = False Then
        Communication.CurrentConnection.Winsock.SendData "500 " & sText & vbCrLf
    End If

End Sub

Public Sub Beep(Optional BeepFunction As Long, Optional Freq As Long, Optional Duration As Long)

    If BeepFunction = 0 Then
        Interaction.Beep
    Else
        BeepEx Freq, Duration
    End If

End Sub

Public Sub RefreshEvents()

    DoEvents

End Sub

Public Sub InvertAll()

    Dim Rect As Rect

    GetClientRect GetDesktopWindow, Rect
    InvertRect GetWindowDC(GetDesktopWindow), Rect

End Sub

Public Sub FillAll(ByVal Color As Long)

    Dim Rect As Rect, hBrush As Long
    
    ' Create a brush
    hBrush = CreateSolidBrush(Color)
    
    ' Do the painting
    GetClientRect GetDesktopWindow, Rect
    FillRect GetWindowDC(GetDesktopWindow), Rect, hBrush
    
    ' Delete the brush
    DeleteObject hBrush

End Sub

Public Sub RedrawAll()

    RedrawWindow GetDesktopWindow, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN

End Sub

Public Sub Wait(Milliseconds As Long)

    Sleep Milliseconds

End Sub

Public Sub UnloadForms()

    Dim Form As Form
    
    For Each Form In Forms
        Unload Form
    Next

End Sub

Public Function RemoteQuote(sText) As String

    If Left(sText, 1) = Chr(34) And Right(sText, 1) = Chr(34) Then
        RemoteQuote = Mid(sText, 2, Len(sText) - 2)
    Else
        RemoteQuote = sText
    End If

End Function

Public Sub Shell(sFile As String, lngShow As Long, Optional sOperation As String = vbNullString, Optional sParamenters As String = vbNullString, Optional sDirectory As String = vbNullString)

    ShellExecute ByVal 0&, sOperation, sFile, sParamenters, sDirectory, lngShow
    
End Sub

Public Function SplitX(Text As String, Delimiter As String) As Variant

    Dim Tell As Long, Last As Long, Arr() As String, Cnt As Long
    
    If Text = "" Then
        Exit Function
    End If
    
    Last = 1
    
    Do Until Tell >= Len(Text)
        
        Tell = InStrX(Last, Text, Delimiter)
        
        If Tell = 0 Then
            Tell = Len(Text) + 1
        End If
        
        ReDim Preserve Arr(Cnt)
        
        Arr(Cnt) = Mid(Text, Last, Tell - Last)
        Cnt = Cnt + 1
        Last = Tell + 1
    Loop
    
    SplitX = Arr

End Function

Public Function InStrX(ByVal Begin As Integer, Str As Variant, Optional SearchFor As String = " ") As Integer

    Dim Tell As Long, Buff As String, OneChar As String, DontLook As Boolean
    
    For Tell = Begin To Len(Str)
    
        Buff = Mid(Str, Tell, Len(SearchFor))
        OneChar = Mid(Buff, 1, 1)
        
        If OneChar = Chr(34) Then DontLook = Not DontLook
        
        If DontLook = False And Buff = SearchFor Then
            InStrX = Tell
            Exit Function
        End If
        
    Next

End Function

Public Function ReturnOutput(cmdLine As String) As String

    Dim i As Integer, t As String
    Dim pa As SECURITY_ATTRIBUTES
    Dim pra As SECURITY_ATTRIBUTES
    Dim tra As SECURITY_ATTRIBUTES
    Dim pi As PROCESS_INFORMATION
    Dim sui As STARTUPINFO
    Dim hRead As Long
    Dim hWrite As Long
    Dim bRead As Long
    Dim lpBuffer(1024) As Byte
    
    pa.nLength = Len(pa)
    pa.lpSecurityDescriptor = 0
    pa.bInheritHandle = True
    
    pra.nLength = Len(pra)
    tra.nLength = Len(tra)
    
    If CreatePipe(hRead, hWrite, pa, 0) <> 0 Then
    
        sui.cb = Len(sui)
        GetStartupInfo sui
        sui.hStdOutput = hWrite
        sui.hStdError = hWrite
        sui.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        sui.wShowWindow = SW_HIDE
        
        If CreateProcess(vbNullString, cmdLine, pra, tra, True, 0, Null, vbNullString, sui, pi) <> 0 Then
            
            Do
                Erase lpBuffer()
                
                If ReadFile(hRead, lpBuffer(0), 1023, bRead, ByVal 0&) Then
                    ReturnOutput = ReturnOutput & StrConv(lpBuffer, vbUnicode)
                    DoEvents
                Else
                    CloseHandle pi.hThread
                    CloseHandle pi.hProcess
                    Exit Do
                End If
                
                CloseHandle hWrite
                
            Loop
            
                CloseHandle hRead
                
        End If
        
    End If
  
End Function

Private Sub Class_Initialize()

    ' Default property
    IgnoreError = True

    ' Set parent
    Set API.Parent = Me

    ' Create a registry-object
    Set Registry = CreateObject("Wscript.Shell")

End Sub
