VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMessage 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Event Clicked()

Private Sub Form_Click()

    ' Tell the master-class about this event
    RaiseEvent Clicked

End Sub

Public Sub Form_Resize()

    lblMessage.Width = Me.Width - 240
    lblMessage.Height = Me.Height - 240

End Sub
