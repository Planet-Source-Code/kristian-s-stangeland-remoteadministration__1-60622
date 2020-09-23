VERSION 5.00
Begin VB.Form frmCommunication 
   Caption         =   "Communication"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtData 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmCommunication.frx":0000
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   2760
      Width           =   990
   End
   Begin VB.TextBox txtRaw 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2760
      Width           =   3675
   End
End
Attribute VB_Name = "frmCommunication"
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

Public WithEvents RefConnection As clsConnection
Attribute RefConnection.VB_VarHelpID = -1
Public ExitConnection As Boolean

Private Sub cmdSend_Click()

    ' Send the string
    RefConnection.SendData txtRaw.Text & vbCrLf
    
    ' Remove the data inside raw
    txtRaw.Text = ""
    txtRaw.SetFocus
    
    ' Move the selection to the last
    txtData.SelStart = Len(txtData.Text)

End Sub

Private Sub Form_Resize()

    If Me.WindowState <> vbMinimized Then
        txtData.Width = Me.ScaleWidth - txtData.Left - 8
        txtData.Height = Me.ScaleHeight - txtRaw.Height - 24
        txtRaw.Width = txtData.Width - cmdSend.Width
        txtRaw.Top = txtData.Top + txtData.Height + 8
        cmdSend.Left = txtRaw.Width + txtRaw.Left
        cmdSend.Top = txtRaw.Top
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If ExitConnection = False Then
        Cancel = True
        Me.Hide
    End If

End Sub

Private Sub RefConnection_DataArrival(sData As String)

    txtData.Text = Right(txtData.Text, 10000) & sData

End Sub

Private Sub RefConnection_DataSending(sData As String)

    txtData.Text = Right(txtData.Text, 10000) & sData

End Sub
