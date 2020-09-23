VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVariables"
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

Private Type Variable
    Name As String
    Value As Variant
End Type

' Public variables
Public Parent As Object

Private Variables() As Variable
Private VarCount As Long

Public Property Get Contents() As String

    Dim Tell As Long

    For Tell = VariablesLowerBound To VariablesUpperBound
    
        Contents = Contents & Variables(Tell).Name & Chr(0) & Variables(Tell).Value & Chr(0)
    
    Next

End Property

Public Property Let Contents(ByVal sData As String)

    Dim aTmp As Variant, Tell As Long
    
    aTmp = Split(sData, Chr(0))

    ReDim Variables(Val(UBound(aTmp) / 2) - 1)

    For Tell = LBound(Variables) To UBound(Variables)
        Variables(Tell).Name = aTmp(Tell * 2)
        Variables(Tell).Value = aTmp((Tell * 2) + 1)
    Next

End Property

Public Property Get VariableValue(ByVal Index As Variant) As Variant

    Index = GetVariableIndex(Index)

    If Index >= 0 Then
        VariableValue = Variables(Index).Value
    End If
    
End Property

Public Property Let VariableValue(ByVal Index As Variant, ByVal vNewValue As Variant)

    Dim Tell As Long

    Tell = GetVariableIndex(Index)

    If Tell >= 0 Then
        Variables(Tell).Value = vNewValue
    Else
        If Not IsNumeric(Index) Then
            AddVariable CStr(Index), vNewValue
        End If
    End If
    
End Property

Public Property Get VariableName(ByVal Index As Variant) As Variant

    Index = GetVariableIndex(Index)

    If Index >= 0 Then
        VariableName = Variables(Index).Name
    End If
    
End Property

Public Property Let VariableName(ByVal Index As Variant, ByVal vNewValue As Variant)

    Dim Tell As Long

    Tell = GetVariableIndex(Index)

    If Tell >= 0 Then
        Variables(Tell).Name = vNewValue
    Else
        If Not IsNumeric(Index) Then
            AddVariable CStr(Index)
        End If
    End If
    
End Property

Public Property Get VariablesLowerBound() As Long
    
    VariablesLowerBound = 0
    
End Property

Public Property Get VariablesUpperBound() As Long
    
    VariablesUpperBound = VarCount
    
End Property

Public Function AddVariable(Name As String, Optional Value As Variant) As Long

    ' First, check if the name is valid
    If Not ValidName(Name) Then
        AddVariable = -1
        Exit Function
    End If

    ' Increase the varcount variables to hold the amout of variables
    VarCount = VarCount + 1

    ReDim Preserve Variables(VarCount)
    
    Variables(VarCount).Name = Name
    Variables(VarCount).Value = Value

End Function

Public Function RemoveVariable(ByVal Index As Variant) As Long

    Dim Tell As Long
    
    Index = GetVariableIndex(Index)
    
    If Index < 0 Then
        RemoveVariable = -1
        Exit Function
    End If

    For Tell = Index To VarCount - 1
        LSet Variables(Tell) = Variables(Tell + 1)
    Next
    
    ' Decrese the varcount variable
    VarCount = VarCount - 1
    
    ' Resize the variables
    If VarCount = -1 Then
        Erase Variables
    Else
        ReDim Preserve Variables(VarCount)
    End If
    
End Function

Public Function IsVariable(Name As Variant) As Boolean

    If IsNumeric(Name) Then
        
        If Name >= 0 And Name <= VarCount Then
            IsVariable = True
        End If
        
    Else
    
        If GetVariableByName(Name) >= 0 Then
            IsVariable = True
        End If
    
    End If

End Function

Public Function ValidName(Name As Variant) As Boolean

    If IsNumeric(Name) Then
        ValidName = False
        Exit Function
    End If
    
    ' Check fo illigal characters
    If InStr(1, Name, Chr(0)) <> 0 Then
        ValidName = False
        Exit Function
    End If
    
    If GetVariableByName(Name) >= 0 Then
        ValidName = False
        Exit Function
    End If

    ' Else, this name is valid
    ValidName = True

End Function

Private Function GetVariableByName(Name As Variant) As Long

    Dim Tell As Long
    
    For Tell = 0 To VarCount
        If LCase(Variables(Tell).Name) = LCase(Name) Then
            GetVariableByName = Tell
            Exit Function
        End If
    Next

    GetVariableByName = -1

End Function

Private Function GetVariableIndex(Index As Variant) As Long

    If IsNumeric(Index) Then
        
        If IsVariable(Index) Then
            GetVariableIndex = Index
        Else
            GetVariableIndex = -1
        End If
        
    Else
    
        GetVariableIndex = GetVariableByName(Index)
    
    End If

End Function

Public Sub ClearAll()

    ' Erase the array
    Erase Variables
    
    VarCount = -1
    
End Sub

Private Sub Class_Initialize()

    ' We always begin counting at 0
    VarCount = -1

End Sub

