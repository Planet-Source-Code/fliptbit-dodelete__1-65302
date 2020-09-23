Attribute VB_Name = "modRegistry"
'//-----------------------------------------------------------------------------------------------------------------
'//    Project: DODelete.vbp
'//   FileName: modRegistry.bas
'//   Coded By: FliptBit
'//
'// Start Date: 02 MAY 2006
'//
'//
'//    Purpose: Access/Read/Write/Delete windows registry keys
'//
'//      Notes:
'//        - Attained programming ideas & information from:
'//           1.  MSDN Knowledge Base Article ID: Q145679
'//           2.  MSDN Knowledge Base Article ID: Q172274
'//           3.  MSDN Knowledge Base Article ID: Q178755
'//           4.  registry.c (Copyright (c) 1993  Microsoft Corporation)
'//
'//******************************************************************************************************************
'//
'//    License:
'//
'//    Copyright (C) 2006  John R. Reid IV (aka FliptBit)
'//
'//    This program is free software; you can redistribute it and/or modify
'//    it under the terms of the GNU General Public License as published by
'//    the Free Software Foundation; either version 2 of the License, or
'//    (at your option) any later version.
'//
'//    This program is distributed in the hope that it will be useful,
'//    but WITHOUT ANY WARRANTY; without even the implied warranty of
'//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'//    GNU General Public License for more details.
'//
'//    You should have received a copy of the GNU General Public License
'//    along with this program; if not, write to the Free Software
'//    Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
'//
'//-----------------------------------------------------------------------------------------------------------------

Option Explicit


'// Windows API Functions
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

'// Registry Error Constants
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234&

'// Registry API Constant Declarations
Private Const REG_SZ = 1&
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL

Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Const KEY_EXECUTE = KEY_READ

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

Private hKey              As Long
Private MainKeyHandle     As Long
Private retval            As Long
Private strBuffer         As String

'//Set the path of the registry keys (saves typing)
Private Const REGISTRY_ROOT = "HKEY_LOCAL_MACHINE\Software\FliptBit Technologies\DODelete"

'//Display registry errors?
Private Const DisplayRegError = False
'
'

Public Function SetStringValue(ByVal sKey As String, ByVal sKeyName As String, ByVal KeyValue As String)
  '//Set the string value of a key.  Returns True if successful or False for error.
  
  '//Set the default state to return
  SetStringValue = False
  
  '//Get the Main Key Handle
  Call ParseKey(sKey, MainKeyHandle)
  If MainKeyHandle Then
    retval = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_WRITE, hKey)
    If retval = ERROR_SUCCESS Then
      retval = RegSetValueEx(hKey, sKeyName, 0, REG_SZ, ByVal KeyValue, Len(KeyValue))
      If Not retval = ERROR_SUCCESS Then
        If DisplayRegError = True Then
          ShowMessage "Registry Error: " & Str$(GetErrorMsg(retval)) & vbCrLf & "Could not set the string value.", "SetStringValue()"
        End If
      Else
        SetStringValue = True
      End If
      retval = RegCloseKey(hKey)
    Else
      If DisplayRegError = True Then
        ShowMessage "Registry Error: " & Str$(GetErrorMsg(retval)) & vbCrLf & "Could open registry key.", "SetStringValue()"
      End If
    End If
  End If
End Function

Public Function GetStringValue(ByVal sKey As String, ByVal sKeyName As String)
  '//Get the string value from a registry key.  Returns value of key or "ERROR"
  
  Dim lBufferSize As Long
  
  lBufferSize = 0
  strBuffer = ""
  
  Call ParseKey(sKey, MainKeyHandle)
  If MainKeyHandle Then
    retval = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)
    If retval = ERROR_SUCCESS Then
      strBuffer = Space(255)
      lBufferSize = Len(strBuffer)
      retval = RegQueryValueEx(hKey, sKeyName, _
        0, REG_SZ, strBuffer, lBufferSize)
      If retval = ERROR_SUCCESS Then
        retval = RegCloseKey(hKey)
        strBuffer = Trim(strBuffer)
        If strBuffer = "" Then
          GetStringValue = "ERROR"
        Else
          GetStringValue = Left(strBuffer, lBufferSize - 1)
        End If
      Else
        '//The VALUE COULD NOT be retreived
        GetStringValue = "ERROR"
        If DisplayRegError = True Then
        ShowMessage "Registry Error: " & Str$(GetErrorMsg(retval)) & vbCrLf & "Could not retrieve the string value", "GetStringValue()"
        End If
      End If
    Else
      '//The KEY COULD NOT be opened
      GetStringValue = "ERROR"
      If DisplayRegError = True Then
        ShowMessage "Registry Error: " & Str$(GetErrorMsg(retval)) & vbCrLf & "Could not open the registry key", "GetStringValue()"
      End If
    End If
  End If
End Function

Public Function CreateKey(ByVal sKey As String)
  '//Creates a registry key
  CreateKey = False
  Call ParseKey(sKey, MainKeyHandle)
  If MainKeyHandle Then
    retval = RegCreateKey(MainKeyHandle, sKey, hKey)
    If retval = ERROR_SUCCESS Then
      retval = RegCloseKey(hKey)
      CreateKey = True
    End If
  End If
End Function

Public Function DeleteKey(ByVal KeyName As String)
  DeleteKey = False
  Call ParseKey(KeyName, MainKeyHandle)
  If MainKeyHandle Then
    retval = RegDeleteKey(MainKeyHandle, KeyName)
    If (retval <> ERROR_SUCCESS) Then
      If DisplayRegError = True Then
        ShowMessage "Registry Error: " & Str$(GetErrorMsg(retval)) & vbCrLf & "Could not delete the registry key.", "DeleteKey()"
      End If
    Else
      DeleteKey = True
    End If
  End If
End Function

Public Function DeleteKeyValue(ByVal sKeyName As String, ByVal sValueName As String)
  DeleteKeyValue = False
  
  Call ParseKey(sKeyName, MainKeyHandle)

  If MainKeyHandle Then
    retval = RegOpenKeyEx(MainKeyHandle, sKeyName, 0, KEY_WRITE, hKey)
    If (retval = ERROR_SUCCESS) Then
      retval = RegDeleteValue(hKey, sValueName)
      If (retval <> ERROR_SUCCESS) Then
        If DisplayRegError = True Then
          ShowMessage "Registry Error: " & Str$(GetErrorMsg(retval)) & vbCrLf & "Could not delete the registry value.", "DeleteKeyValue()"
        End If
      Else
        DeleteKeyValue = True
      End If
      retval = RegCloseKey(hKey)
    End If
  End If
End Function

Public Function DeleteAllKeySubItems()
  DeleteAllKeySubItems = False
End Function

Public Function KeyExist(ByVal sKey As String)
  Call ParseKey(sKey, MainKeyHandle)
  If MainKeyHandle Then
    retval = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)
    If retval = ERROR_SUCCESS Then
      KeyExist = True
    Else
      KeyExist = False
    End If
  End If
End Function

Public Function KeyValueExist(ByVal sKey As String, ByVal sKeyName As String)
  Dim lActualType As Long
  Dim lSize As Long
  Dim sTmp As String

  Call ParseKey(sKey, MainKeyHandle)
  If MainKeyHandle Then
    retval = RegOpenKeyEx(MainKeyHandle, sKey, 0, KEY_READ, hKey)
    If (retval = ERROR_SUCCESS) Then
      retval = RegQueryValueEx(hKey, ByVal sKeyName, 0&, lActualType, sTmp, lSize)
      If (retval = ERROR_SUCCESS) Then
        KeyValueExist = True
      Else
        KeyValueExist = False
      End If
    End If
  End If
End Function

Private Sub ParseKey(KeyName As String, Keyhandle As Long)
   '//Return if "\" is contained in the Keyname
  retval = InStr(KeyName, "\")
  
  '//If the is a "\" at the end of the Keyname then
  If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then
    ShowMessage "Incorrect Format: " & KeyName, "ParseKey()"
    Exit Sub
  ElseIf retval = 0 Then
    '//If the Keyname contains no "\"
    Keyhandle = GetMainKeyHandle(KeyName)
    KeyName = ""
  Else
    '//Otherwise, Keyname contains "\" --  seperate the Keyname
    Keyhandle = GetMainKeyHandle(Left(KeyName, retval - 1))
    KeyName = Right(KeyName, Len(KeyName) - retval)
  End If
End Sub

Private Function GetMainKeyHandle(MainKeyName As String) As Long
  MainKeyName = UCase$(MainKeyName)
  Select Case MainKeyName
    Case "HKEY_CLASSES_ROOT":       GetMainKeyHandle = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER":       GetMainKeyHandle = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE":      GetMainKeyHandle = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS":              GetMainKeyHandle = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA":   GetMainKeyHandle = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG":     GetMainKeyHandle = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA":           GetMainKeyHandle = HKEY_DYN_DATA
  End Select
End Function

Private Function GetErrorMsg(lErrorCode As Long) As String
  Select Case lErrorCode
    Case 1009, 1015:  GetErrorMsg = "Registry Database is corrupt"
    Case 2, 1010:     GetErrorMsg = "Bad Key Name"
    Case 1011:        GetErrorMsg = "Can't Open Key"
    Case 4, 1012:     GetErrorMsg = "Can't Read Key"
    Case 5:           GetErrorMsg = "Access to this key is denied"
    Case 1013:        GetErrorMsg = "Can't Write Key"
    Case 8, 14:       GetErrorMsg = "Out of memory"
    Case 87:          GetErrorMsg = "Invalid Parameter"
    Case 234:         GetErrorMsg = "There is more data than the buffer has been allocated to hold."
    Case Else:        GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
  End Select
End Function

Public Sub RegCreateRoot()
  CreateKey REGISTRY_ROOT
End Sub

Public Sub RegDeleteRoot()
  DeleteKey REGISTRY_ROOT
End Sub

Public Function RegReadOverwritePattern() As String
  Dim rResult As String
  rResult = GetStringValue(REGISTRY_ROOT, "Pattern")
  RegReadOverwritePattern = rResult
End Function

Public Sub RegWriteOverwritePattern(value As String)
  SetStringValue REGISTRY_ROOT, "Pattern", value
End Sub

Public Function RegReadRenameFiles() As String
  Dim rResult As String
  rResult = GetStringValue(REGISTRY_ROOT, "Rename")
  RegReadRenameFiles = rResult
End Function

Public Sub RegWriteRenameFiles(value As String)
  SetStringValue REGISTRY_ROOT, "Rename", value
End Sub
