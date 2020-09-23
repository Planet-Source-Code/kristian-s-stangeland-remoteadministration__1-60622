VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
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

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3

Public Property Get Name() As String

    Name = "Registry"
    
End Property

Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long, strData As Integer
    
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    
    If lResult = 0 Then
        If lValueType = REG_SZ Or lValueType = REG_EXPAND_SZ Then

            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            
            If lResult = 0 Then
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If

        ElseIf lValueType = REG_BINARY Then
        
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
            
        End If
        
    End If
    
End Function

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    
    Dim Ret As Long

    RegOpenKey hKey, strPath, Ret
    GetString = RegQueryStringValue(Ret, strValue)
    RegCloseKey Ret
    
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    
    Dim Ret As Long
    
    RegCreateKey hKey, strPath, Ret
    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    RegCloseKey Ret
    
End Sub

Public Sub SaveStringLong(hKey As Long, strPath As String, strValue As String, strData As String)
    
    Dim Ret As Long

    RegCreateKey hKey, strPath, Ret
    RegSetValueEx Ret, strValue, 0, REG_BINARY, CByte(strData), 4
    RegCloseKey Ret
    
End Sub

Public Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    
    Dim Ret As Long

    RegCreateKey hKey, strPath, Ret
    RegDeleteValue Ret, strValue
    RegCloseKey Ret
    
End Sub


