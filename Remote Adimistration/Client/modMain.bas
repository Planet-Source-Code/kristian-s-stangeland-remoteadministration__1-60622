Attribute VB_Name = "modGlobal"
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

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const SW_SHOWNORMAL = 1
Public Const BITSPIXEL = 12
Public Const ICC_USEREX_CLASSES = &H200

Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

' Constant for file-sending
Public Const ChunckSize = 2048

' If we're supposed to load true colour icons
Public m_bTrueColour As Boolean

' Plugins
Public Plugins As New clsPlugins

Public Function TempFile(Optional sPrefix As String = "RAC") As String

    Dim sFile As String, sPath As String

    sPath = String(100, 0)
    sFile = String(260, 0)

    GetTempPath 100, sPath
    GetTempFileName sPath, sPrefix, 0, sFile

    TempFile = Left(sFile, InStr(1, sFile, Chr(0)) - 1)

End Function

Public Function InitCommonControlsVB() As Boolean

   On Error Resume Next
   
   Dim iccex As tagInitCommonControlsEx
   
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   
   On Error GoTo 0
   
End Function

Public Sub Main()

    If InitCommonControlsVB = False Then
        MsgBox "Could not load comctl32.dll", vbCritical, "Error"
        End
    End If
    
    If GetDeviceCaps(GetDC(0), BITSPIXEL) > 8 Then
       m_bTrueColour = True
    End If
    
    ' Load plugins
    Plugins.LoadPlugins
    
    ' Show main form
    frmLogon.Show

End Sub

Public Sub WriteFile(sFile As String, sData As String, Optional lngPosition As Long)

    Dim Free As Long
    
    Free = FreeFile
    
    Open sFile For Binary As #Free
        
        If lngPosition < 1 Then
            lngPosition = LOF(Free) + 1
        End If
    
        Put #Free, lngPosition, sData
    Close #Free

End Sub

Public Function GetFileName(sFile As String) As String

    GetFileName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))

End Function

Public Function GetFileBase(sFile As String) As String

    Dim sTemp As String
    
    ' Get the file name
    sTemp = GetFileName(sFile)

    ' Return the file base
    GetFileBase = Left(sTemp, InStr(1, sTemp, ".") - 1)

End Function

Public Function GetFileExtension(sFile As String) As String

    GetFileExtension = Right(sFile, Len(sFile) - InStrRev(sFile, "."))

End Function

Public Function ValidPath(sPath As String) As String

    ValidPath = sPath & IIf(Right(sPath, 1) = "\", "", "\")

End Function

Public Function IsDir(sPath As String) As Boolean

    IsDir = CBool(InStr(1, sPath, ".") <= 0)

End Function

Public Function RemoveLetters(sText As String) As String

    Dim Tell As Long, Char As String
    
    For Tell = 1 To Len(sText)
    
        ' Retrive the char from the string
        Char = Mid(sText, Tell, 1)
    
        ' Check and see if it's a number - if so, add it to the output
        If IsNumeric(Char) Then
            RemoveLetters = RemoveLetters & Char
        End If
        
    Next

End Function

' Old function. I didn't bother rewriting it.
Public Function KeyName(KeyCode As Integer, Shift As Long, Caps_Lock As Long) As String

    Dim sTemp As String
    
    Select Case KeyCode
    Case 65 To 90: sTemp = LCase(Chr(KeyCode))
    Case 221: sTemp = "å"
    Case 192: sTemp = "ø"
    Case 222: sTemp = "æ"
    Case 48 To 57: sTemp = KeyCode - 48
    Case 96 To 105 And Shift >= 0: sTemp = KeyCode - 96
    Case vbKeyAdd: sTemp = "+"
    Case 13: sTemp = vbNewLine
    Case vbKeyBack: sTemp = " (BACK) "
    Case vbKeyLeft: sTemp = " (LEFT) "
    Case vbKeyRight: sTemp = " (RIGHT) "
    Case vbKeyInsert: sTemp = " (INSERT) "
    Case vbKeyHome: sTemp = " (HOME) "
    Case vbKeyPageDown: sTemp = " (PAGE DOWN) "
    Case vbKeyUp: sTemp = " (UP) "
    Case vbKeyDown: sTemp = " (DOWN) "
    Case vbKeyEnd: sTemp = " ((END))"
    Case vbKeyDelete: sTemp = " (DELETE) "
    Case vbKeyPageUp: sTemp = " (PAGE UP) "
    Case vbKeyScrollLock: sTemp = " (SCROLL LOCK) "
    Case vbKeyPrint: sTemp = " (PRINT) "
    Case vbKeyPause: sTemp = " (PAUSE) "
    Case 112 To 123: sTemp = " (F" & KeyCode - 111 & ") "
    Case vbKeyShift: sTemp = " (SHIFT) "
    Case vbKeyControl: sTemp = " (CONTROL) "
    Case 18: sTemp = "(ALT)"
    Case 9: sTemp = "(TAB)"
    Case vbKeyEscape: sTemp = " (ESCAPE) "
    Case 220: sTemp = " (FRACTION 1/2) "
    Case 32: sTemp = " "
    Case 92: sTemp = " (RIGHT WINDOW BUTTON) "
    Case 91: sTemp = " (LEFT WINDOW BUTTON) "
    Case 191: sTemp = "*"
    Case 189: sTemp = "-"
    Case 188: sTemp = ","
    Case 219: sTemp = "\"
    Case 187: sTemp = "?"
    Case 193: sTemp = " (MENU) "
    Case 190: sTemp = "."
    Case 144: sTemp = " (NUM LOCK) "
    Case 111: sTemp = " (DIVISION SIGN) "
    Case 106: sTemp = " (MULTIPLICATION SIGN) "
    Case 109: sTemp = " (MINUS SIGN) "
    Case 110: sTemp = " (RIGHT DELETE) "
    Case 226: sTemp = "<"
    Case 186: sTemp = "¨"
    Case 93: sTemp = " (MENY) "
    Case Else: sTemp = Chr(KeyCode)
    End Select
    
    If Shift < 0 Then
        Select Case KeyCode
        Case 65 To 90: sTemp = UCase(Chr(KeyCode))
        Case 221: sTemp = "Å"
        Case 192: sTemp = "Ø"
        Case 222: sTemp = "Æ"
        Case vbKey0: sTemp = "="
        Case vbKey1: sTemp = "!"
        Case vbKey2: sTemp = Chr(34)
        Case vbKey3: sTemp = "#"
        Case vbKey4: sTemp = "¤"
        Case vbKey5: sTemp = "%"
        Case vbKey6: sTemp = "&"
        Case vbKey7: sTemp = "/"
        Case vbKey8: sTemp = "("
        Case vbKey9: sTemp = ")"
        Case vbKeyAdd: sTemp = "?"
        Case 219: sTemp = "`"
        Case 186: sTemp = "^"
        Case 190: sTemp = ":"
        Case 188: sTemp = ";"
        Case 226: sTemp = ">"
        End Select
    End If
    
    If Caps_Lock < 0 Then
        Select Case KeyCode
        Case 65 To 90: sTemp = UCase(Chr(KeyCode))
        Case 221: sTemp = "Å"
        Case 192: sTemp = "Ø"
        Case 222: sTemp = "Æ"
        End Select
    End If
    
    KeyName = sTemp

End Function
