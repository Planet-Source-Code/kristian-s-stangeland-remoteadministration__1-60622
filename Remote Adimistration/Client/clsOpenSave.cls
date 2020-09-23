VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpenSave"
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

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Enum OFN_FLAGS
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
End Enum

Private SaveFileDialog As OPENFILENAME
Private OpenFileDialog As OPENFILENAME

Private rv As Long
Private sv As Long
Private lFilter As String
Private lMaxFile As Long
Private lFlags As Long
Private lFile As String

Public Sub OpenFile(hwnd As Long, Caption As String)
   
    With OpenFileDialog
     .lStructSize = Len(OpenFileDialog)
     .hwndOwner = hwnd
     .hInstance = App.hInstance
     .lpstrFilter = lFilter
     .lpstrFile = Space$(lMaxFile - 1)
     .nMaxFile = lMaxFile
     .lpstrFileTitle = Space$(lMaxFile - 1)
     .nMaxFileTitle = 255
     .lpstrInitialDir = App.Path
     .lpstrTitle = Caption
     .Flags = lFlags
    End With
    
    rv& = GetOpenFileName(OpenFileDialog)
    If (rv&) Then lFile = Trim$(OpenFileDialog.lpstrFile) Else lFile = ""

End Sub

Public Sub SaveFile(hwnd As Long, Caption As String)
   
    With SaveFileDialog
     .lStructSize = Len(OpenFileDialog)
     .hwndOwner = hwnd
     .hInstance = App.hInstance
     .lpstrFilter = lFilter
     .lpstrFile = Space$(lMaxFile - 1)
     .nMaxFile = lMaxFile
     .lpstrFileTitle = Space$(lMaxFile - 1)
     .nMaxFileTitle = 255
     .lpstrInitialDir = App.Path
     .lpstrTitle = Caption
     .Flags = lFlags
    End With
    
    sv& = GetSaveFileName(SaveFileDialog)
    If (sv&) Then lFile = Trim$(SaveFileDialog.lpstrFile) Else lFile = ""

End Sub

Public Property Get Name() As String

    Name = "OpenSave"
    
End Property

Public Property Get MaxFile() As Long

    MaxFile = lMaxFile

End Property

Public Property Let MaxFile(ByVal vNewValue As Long)

    lMaxFile = vNewValue

End Property

Public Property Get Filter() As String

    Filter = lFilter

End Property

Public Property Let Filter(ByVal vNewValue As String)

    lFilter = vNewValue

End Property

Public Property Get Flags() As OFN_FLAGS

    Flags = lFlags

End Property

Public Property Let Flags(ByVal vNewValue As OFN_FLAGS)

    lFlags = vNewValue

End Property

Public Property Get File() As String

    File = lFile

End Property

Public Property Let File(ByVal vNewValue As String)

    lFile = vNewValue

End Property

Private Sub Class_Initialize()

    lMaxFile = 255

End Sub
