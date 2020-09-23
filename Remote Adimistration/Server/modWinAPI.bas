Attribute VB_Name = "modWinAPI"
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

Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetVersion Lib "kernel32" () As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

Declare Function BeepEx Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long

Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetTickCount Lib "kernel32" () As Long

Declare Function CreateTimerQueue Lib "kernel32.dll" () As Long
Declare Function CreateTimerQueueTimer Lib "kernel32.dll" (ByRef phNewTimer As Long, ByVal TimerQueue As Long, ByVal Callback As Long, ByVal Parameter As Long, ByVal DueTime As Long, ByVal Period As Long, ByVal flags As Long) As Long
Declare Function DeleteTimerQueue Lib "kernel32.dll" (ByVal TimerQueue As Long) As Long
Declare Function DeleteTimerQueueTimer Lib "kernel32.dll" (ByVal TimerQueue As Long, ByVal Timer As Long, ByVal CompletionEvent As Long) As Long

Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As Rect) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)

Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100
Public Const RDW_ALLCHILDREN = &H80

Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

Public Const DRIVE_UNKNOWN = 0
Public Const DRIVE_ABSENT = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Public Const ERROR_BAD_DEVICE = 1200&
Public Const ERROR_CONNECTION_UNAVAIL = 1201&
Public Const ERROR_EXTENDED_ERROR = 1208&
Public Const ERROR_MORE_DATA = 234
Public Const ERROR_NOT_SUPPORTED = 50&
Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&
Public Const ERROR_NO_NETWORK = 1222&
Public Const ERROR_NOT_CONNECTED = 2250&
Public Const NO_ERROR = 0

Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Const ANYSIZE_ARRAY = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const KEYEVENTF_KEYUP = &H2
Public Const INPUT_MOUSE = 0
Public Const INPUT_KEYBOARD = 1
Public Const INPUT_HARDWARE = 2

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const SW_SHOWNORMAL = 1
Public Const SW_HIDE = 0

Public Const SYNCHRONIZE = &H100000
Public Const INFINITE = &HFFFF

Public Const STARTF_USESHOWWINDOW = &H1
Public Const STARTF_USESTDHANDLES = &H100

Enum EWX
    EWX_LOGOFF = 0
    EWX_SHUTDOWN = 1
    EWX_REBOOT = 2
    EWX_FORCE = 4
End Enum

Type POINTAPI
    X As Long
    Y As Long
End Type

Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type MOUSEINPUT
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Type HARDWAREINPUT
    uMsg As Long
    wParamL As Integer
    wParamH As Integer
End Type

Type GENERALINPUT
    dwType As Long
    xi(0 To 23) As Byte
End Type

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Type LUID
    LowPart As Long
    HighPart As Long
End Type

Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type

Type STARTUPINFO
  cb As Long
  lpReserved As Long
  lpDesktop As Long
  lpTitle As Long
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Byte
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Public Sub EnableShutDown()

    Dim hProc As Long
    Dim hToken As Long
    Dim mLUID As LUID
    Dim mPriv As TOKEN_PRIVILEGES
    Dim mNewPriv As TOKEN_PRIVILEGES
    
    hProc = GetCurrentProcess()
    OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeShutdownPrivilege", mLUID
    mPriv.PrivilegeCount = 1
    mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    mPriv.Privileges(0).pLuid = mLUID

    AdjustTokenPrivileges hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount)

End Sub

Public Sub MakeMeService()

    Dim pid As Long, reserv As Long

    pid = GetCurrentProcessId()
    reserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
    
End Sub

Public Sub UnMakeMeService()

    Dim pid As Long, reserv As Long

    pid = GetCurrentProcessId()
    reserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
    
End Sub
