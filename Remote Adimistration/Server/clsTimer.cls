VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
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

Private Type Timer
    Code As String
    Interval As Long
    Enabled As Boolean
    Handle As Long
    Connection As clsConnection
End Type

Private Timer() As Timer
Private hQueue As Long

Public Property Get TimerCount() As Long

    TimerCount = UBound(Timer)

End Property

Public Property Get TimerHandle(ByVal Index As Long) As Boolean

    If IsTimer(Index) Then
        TimerHandle = Timer(Index).Handle
    End If

End Property

Public Property Get TimerEnabled(ByVal Index As Long) As Boolean

    If IsTimer(Index) Then
        TimerEnabled = Timer(Index).Enabled
    End If

End Property

Public Property Get TimerCode(ByVal Index As Long) As String

    If IsTimer(Index) Then
        TimerCode = Timer(Index).Code
    End If

End Property

Public Property Let TimerCode(ByVal Index As Long, ByVal vNewValue As String)

    If IsTimer(Index) Then
        Timer(Index).Code = vNewValue
    End If

End Property

Public Property Get TimerInterval(ByVal Index As Long) As Long

    If IsTimer(Index) Then
        TimerInterval = Timer(Index).Interval
    End If

End Property

Public Property Let TimerInterval(ByVal Index As Long, ByVal vNewValue As Long)

    If IsTimer(Index) Then
        Timer(Index).Interval = vNewValue
    End If

End Property

Public Function AddTimer(sCode As String, lngInterval As Long, Connection As clsConnection) As Long

    Dim Tell As Long
    
    For Tell = LBound(Timer) To UBound(Timer)
        If Timer(Tell).Enabled = False Then
            Exit For
        End If
    Next
    
    If Tell > UBound(Timer) Then
        ReDim Preserve Timer(Tell)
    End If
    
    Set Timer(Tell).Connection = Connection
    Timer(Tell).Code = sCode
    Timer(Tell).Interval = lngInterval
    Timer(Tell).Enabled = True
    
    ' Create the timer
    If CreateTimerQueueTimer(Timer(Tell).Handle, hQueue, AddressOf TimerCallBack, Tell, 0, lngInterval, 0) = 0 Then
        AddTimer = -1
        Exit Function
    End If
    
    ' Return the index
    AddTimer = Tell

End Function

Public Sub RemoveTimer(ByVal Index As Long)

    If IsTimer(Index) Then
        If Timer(Index).Enabled Then
        
            ' Disable the timer-object
            Timer(Index).Enabled = False
            
            ' Clean up
            DeleteTimerQueueTimer hQueue, Timer(Index).Handle, ByVal 0&
            
        End If
    End If
    
End Sub

Public Sub RemoveAll()

    Dim Tell As Long
    
    For Tell = LBound(Timer) To UBound(Timer)
        If Timer(Tell).Enabled Then
            RemoveTimer Tell
        End If
    Next

End Sub

Public Sub RemoveByConnection(ConnectionIndex As Long)

    On Error Resume Next
    Dim Tell As Long

    For Tell = LBound(Timer) To UBound(Timer)
        If Timer(Tell).Connection.Index = ConnectionIndex And Timer(Tell).Enabled Then
            RemoveTimer Tell
        End If
    Next
    
End Sub

Public Sub FireTimer(ByVal Index As Long)

    ' Run the timer
    If IsTimer(Index) Then
        Timer(Index).Connection.RunCode TimerCode(Index)
    End If

End Sub

Public Function IsTimer(ByVal Index As Long) As Boolean

    IsTimer = CBool(Index >= LBound(Timer) And Index <= UBound(Timer))

End Function

Private Sub Class_Initialize()

    ' Start with one timer-object
    ReDim Timer(0)

    ' Make timers possible
    hQueue = CreateTimerQueue()
    
End Sub

Private Sub Class_Terminate()

    ' Remove all timers
    RemoveAll

    ' Clean up
    DeleteTimerQueue hQueue

End Sub
