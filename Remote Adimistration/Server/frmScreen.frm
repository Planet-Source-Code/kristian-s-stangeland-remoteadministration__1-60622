VERSION 5.00
Begin VB.Form frmScreen 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   Icon            =   "frmScreen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmScreen"
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

Public Sub FillDesktop()

    PaintDesktop Me.hdc

End Sub

Public Property Let TopMost(ByVal bValue As Boolean)

    SetWindowPos Me.hwnd, IIf(bValue, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Property
