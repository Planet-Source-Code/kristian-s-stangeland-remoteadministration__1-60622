VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessage"
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

' Public variables
Public TimeOut As Long
Public StopAtClick As Boolean
Public StopAtEsc As Boolean

' This class' form
Private WithEvents oMessage As frmMessage
Attribute oMessage.VB_VarHelpID = -1

' Internal variables
Private sMessage As String
Private sTitle As String
Private lX As Long
Private lY As Long
Private lWidth As Long
Private lHeight As Long

Public Property Get Message() As String

    Message = sMessage

End Property

Public Property Let Message(ByVal vNewValue As String)

    sMessage = vNewValue
    
    If Not oMessage Is Nothing Then
        oMessage.lblMessage.Caption = sMessage
    End If

End Property

Public Property Get Title() As String

    Title = sTitle

End Property

Public Property Let Title(ByVal vNewValue As String)

    sTitle = vNewValue
    
    If Not oMessage Is Nothing Then
        oMessage.Caption = sTitle
    End If

End Property

Public Property Get X() As Long

    X = lX

End Property

Public Property Let X(ByVal vNewValue As Long)

    lX = vNewValue
    
    If Not oMessage Is Nothing Then
        oMessage.Left = lX
    End If

End Property

Public Property Get Y() As Long

    Y = lY

End Property

Public Property Let Y(ByVal vNewValue As Long)

    lY = vNewValue
    
    If Not oMessage Is Nothing Then
        oMessage.Top = lY
    End If

End Property

Public Property Get Width() As Long

    Width = oMessage.ScaleX(lWidth, vbTwips, vbPixels)

End Property

Public Property Let Width(ByVal vNewValue As Long)

    lWidth = oMessage.ScaleX(vNewValue, vbPixels, vbTwips)
    
    If Not oMessage Is Nothing Then
        If oMessage.Visible And lWidth > 280 Then
            oMessage.Width = lWidth
            oMessage.Form_Resize
        End If
    End If

End Property

Public Property Get Height() As Long

    Height = oMessage.ScaleY(lHeight, vbTwips, vbPixels)

End Property

Public Property Let Height(ByVal vNewValue As Long)

    lHeight = oMessage.ScaleY(vNewValue, vbPixels, vbTwips)
    
    If Not oMessage Is Nothing Then
        If oMessage.Visible And lHeight > 280 Then
            oMessage.Height = lHeight
            oMessage.Form_Resize
        End If
    End If

End Property

Public Property Get MessageForm() As Form

    Set MessageForm = oMessage

End Property

Public Sub HideMessageForm()

    oMessage.Hide

End Sub

Public Sub ShowMessageForm(Optional Async As Boolean)

    Dim StartTime As Long

    ' Set the message and position of the form
    oMessage.lblMessage.Caption = sMessage
    oMessage.Left = lX
    oMessage.Top = lY
    oMessage.Width = IIf(lWidth > 280, lWidth, 280)
    oMessage.Height = IIf(lHeight > 280, lHeight, 280)
    
    ' Invoke the resize-prosedure
    oMessage.Form_Resize
    
    ' Show the form
    oMessage.Show

    If Async = False Then
        ' Get the start-up time
        StartTime = GetTickCount
        
        Do Until oMessage.Visible = False Or ((GetTickCount - StartTime > TimeOut) And TimeOut > 0)
            
            If StopAtEsc Then
                If GetKeyState(vbKeyEscape) < 0 Then
                    Exit Sub
                End If
            End If
        
            Sleep 10
            DoEvents
        Loop
    End If

End Sub

Public Function ShowMessageBox(Optional Buttons As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult

    ShowMessageBox = MsgBox(sMessage, Buttons, sTitle)

End Function

Public Function Clone() As clsMessage

    Set Clone = New clsMessage
    
    Clone.X = lX
    Clone.Y = lY
    Clone.Width = lWidth
    Clone.Height = lHeight
    Clone.Message = sMessage
    Clone.Title = sTitle
    Clone.StopAtClick = StopAtClick
    Clone.StopAtEsc = StopAtEsc

End Function

Private Sub Class_Initialize()

    ' Create the form
    Set oMessage = New frmMessage

    ' Everything is twips
    lWidth = 3765
    lHeight = 1185

End Sub

Private Sub Class_Terminate()

    ' Destroy the form
    Set oMessage = Nothing

End Sub

Private Sub oMessage_Clicked()

    If StopAtClick Then
        oMessage.Hide
    End If

End Sub
