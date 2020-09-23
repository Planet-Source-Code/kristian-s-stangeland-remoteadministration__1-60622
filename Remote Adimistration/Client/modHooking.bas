Attribute VB_Name = "modHooking"
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

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Const WM_PAINT = &HF
Public Const GWL_WNDPROC = (-4)

Public Sub HookControl(hwnd As Long, Control As Object)
    
    SetProp hwnd, "ObjPtr", ObjPtr(Control)
    SetProp hwnd, "PrevProc", SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Public Sub UnHookControl(hwnd As Long)

    Dim PrevProc As Long
    
    PrevProc = GetProp(hwnd, "PrevProc")
    
    If PrevProc <> 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, PrevProc
        RemoveProp hwnd, "PrevProc"
        RemoveProp hwnd, "ObjPtr"
    End If
    
End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
    Dim PrevProc As Long, Control As Object, lSrc As Long
    
    PrevProc = GetProp(hwnd, "PrevProc")
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    
    If uMsg = WM_PAINT Then
    
        lSrc = GetProp(hwnd, "ObjPtr")
        
        If lSrc <> 0 Then
            CopyMemory ByVal VarPtr(Control), lSrc, 4
            Control.Parent.UpdateColors
        End If
        
    End If
    
End Function
