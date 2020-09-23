VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add user"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   990
      Width           =   2055
   End
   Begin VB.Label lblUserName 
      Caption         =   "&User name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   510
      Width           =   2055
   End
End
Attribute VB_Name = "frmUser"
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

Private Sub cmdCancel_Click()

    Me.Tag = "CANCEL"
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Tag = "OK"
    Me.Hide

End Sub

