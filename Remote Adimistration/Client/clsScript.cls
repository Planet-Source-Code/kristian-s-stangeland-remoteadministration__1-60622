VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugins"
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

Public Plugins As Collection

Public Property Get MainForm() As Form

    Set MainForm = frmMain
    
End Property

Public Property Get CollectionForms() As Object

    Set CollectionForms = Forms
    
End Property

Public Property Get Name() As String

    Name = "Plugins"
    
End Property

Public Sub LoadPlugins()

    Dim sFile As Variant, Plugin As Object, strClassName As String
    
    For Each sFile In GetFolderList(App.Path & "\Plugins\", "dll")
    
        Select Case GetFileExtension(CStr(sFile))
        Case "dll"
        
            ' Clear earlier errors
            Err.Clear
            
            ' How the class is registered in the registry
            strClassName = GetFileBase(CStr(sFile)) & ".PluginMain"
            
            ' Try to create the object
            Set Plugin = CreateObject(strClassName)
    
            If Err = 429 Then ' ERROR: ActiveX component can't create object
                ' Try to register the object
                Shell "regsvr32 " & Chr(34) & ValidPath(App.Path) & "Plugins\" & CStr(sFile) & Chr(34)
                
                ' Load the plugin again
                Set Plugin = CreateObject(strClassName)
            End If
    
            ' Add plugin
            Plugins.Add Plugin, Plugin.Name
    
            ' Initialize plugin
            Plugin.Initialize Me
        
        End Select
        
    Next

End Sub

Public Function GetFolderList(Path As String, ByVal Extention As Variant) As Collection

    Dim sFile As String
    
    ' Create a new collection
    Set GetFolderList = New Collection
    
    sFile = Dir(Path)
    
    ' Loop through all files
    Do Until sFile = ""
    
        If GetFileExtension(sFile) = Extention Or Extention = "*" Then
            GetFolderList.Add sFile
        End If
    
        ' Go to next file
        sFile = Dir
    Loop

End Function

