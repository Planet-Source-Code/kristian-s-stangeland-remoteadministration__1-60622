VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client - Log on"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "C&onnect"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtComputer 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblUserName 
      Caption         =   "&User name:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   990
      Width           =   2055
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1350
      Width           =   2055
   End
   Begin VB.Label lblComputer 
      Caption         =   "&Computer:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogon"
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

    Unload Me

End Sub

Private Sub cmdConnect_Click()

    Dim Main As New frmMain
    
    ' Tell the user that we're connecting
    cmdConnect.Caption = "Connecting ..."
    cmdConnect.Enabled = False
    
    With Main
        Set .Connection = New clsConnection
        Set .frmCom.RefConnection = .Connection
        
        ' Set username and password
        .Username = txtUsername.Text
        .Password = txtPassword.Text
        
        .State = -1
        .Connection.Connect txtComputer
                
        ' Wait until something has happened
        Do Until .Connection.ConnectError Or .Connection.Connected
            Sleep 10
            DoEvents
        Loop
        
        If .Connection.Connected Then
            ' Show the form and hide this

            If .WaitForState(0, 10000) <> 0 Then
                MsgBox "Error: Timeout", vbCritical, "Error"
            Else
            
                ' Try to connect with username and password
                If txtUsername.Text <> "" Then
                    .Connection.SendData "user " & Chr(34) & txtUsername.Text & Chr(34) & vbCrLf
                    .WaitForData
                End If
                
                If txtPassword.Text <> "" Then
                    .Connection.SendData "pass " & Chr(34) & txtPassword.Text & Chr(34) & vbCrLf
                    .WaitForData
                End If
                
                If Left(.sText, 3) = "530" Then
                    ' We didn't manage to log in
                    Set Main = Nothing
                    MsgBox "Login incorrect", vbCritical, "Error"
                Else
                    .Show
                    Me.Hide
                End If
            End If

        Else
            ' Remote the form
            Set Main = Nothing
        End If
        
    End With
    
    ' Reset the command
    cmdConnect.Caption = "Connect"
    cmdConnect.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim Form As Form
    
    For Each Form In Forms
        Unload Form
    Next

End Sub
