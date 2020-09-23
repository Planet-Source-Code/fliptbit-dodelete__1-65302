Attribute VB_Name = "modDODelete"
'//-----------------------------------------------------------------------------------------------------------------
'//    Project: DODelete.vbp
'//   FileName: mod_DODelete.bas
'//   Coded By: FliptBit
'//
'// Start Date: 13 APR 2006
'//
'//
'//    Purpose: This is my attempt to adhere to the DOD 5220.22-M standard as stated
'//             in the Nation Industrial Security Program Operating Manual (REV Jan 1995).
'//
'//
'//     To Do/Issues:
'//       - Optimize code & variable usage by re-writing module into a class structure.
'//       - Write code to sector-align the array in order to use FILE_FLAG_NO_BUFFERING flag.
'//       - Test to make sure we are actually writing to disk without cache.
'//       - Add code for more/better error handling.
'//
'//
'//      Notes: This module uses Windows API to overwrite the selected file.  This is much
'//             faster than using PRINT statements, and provides the ability to ensure we
'//             are writing straight to disk without using the cache.  The most time consuming
'//             part of the code (aside from GUI stuff) is building the data arrays.
'//
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

'// ----- Define Data Structures -----

'//The FILETIME struct is a 64-bit value representing the number
'//of 100-nanosecond intervals since January 1, 1601.
Private Type FILETIME
  dwLowDateTime     As Long                           '//Specifies the low-order 32 bits of the file time
  dwHighDateTime    As Long                           '//Specifies the high-order 32 bits of the file time
End Type


'//The SYSTEMTIME structure represents a date and time using
'//individual members for the month, day, year, weekday,
'//hour, minute, second, and millisecond.
Private Type SYSTEMTIME
  wYear             As Integer                        '//Specifies the current year
  wMonth            As Integer                        '//Specifies the current month; January = 1, February = 2, and so on
  wDayOfWeek        As Integer                        '//Specifies the current day of the week; Sunday = 0, Monday = 1, and so on
  wDay              As Integer                        '//Specifies the current day of the month
  wHour             As Integer                        '//Specifies the current hour
  wMinute           As Integer                        '//Specifies the current minute
  wSecond           As Integer                        '//Specifies the current second
  wMilliseconds     As Integer                        '//Specifies the current millisecond
End Type

Private Type FILE_INFORMATION
  fName             As String                         '//File Name (Full Path)
  fSize             As Long                           '//File Size (in bytes)
  CreationDate      As Date                           '//File Creation Date
  AccessDate        As Date                           '//Last Access Date
  ModifyDate        As Date                           '//Last Modified Date
End Type


'// The WIN32_FIND_Data Struct to holds data returned from FindFirstFile,FindNextFile
Private Type WIN32_FIND_DATA
  dwFileAttributes      As Long                       '//FILE_ATTRIBUTE_ values as defined in Winnt.h
  ftCreationTime        As FILETIME                   '//specifies when a file or directory is created
  ftLastAccessTime      As FILETIME                   '//specifies when the file is last read from, written to, or for executable files, run
  ftLastWriteTime       As FILETIME                   '//specifies when the file is last written to, truncated, or overwritten
  nFileSizeHigh         As Long                       '//The high-order DWORD value of a file size, in bytes. This value is 0 (zero) unless the file size is greater than MAXDWORD
  nFileSizeLow          As Long                       '//The low-order DWORD value of the file size, in bytes
  dwReserved0           As Long                       '//If the dwFileAttributes member includes the FILE_ATTRIBUTE_REPARSE_POINT attribute, this member specifies the reparse tag
  dwReserved1           As Long                       '//Reserved for future use
  cFileName             As String * 260               '//A null-terminated string that specifies the name of a file
  cAlternateFileName    As String * 14                '//A null-terminated string that specifies an alternative name for a file
End Type


'//*
'// ----- Define Public Constants & Types ------
'//*

Public FILE_INFO  As FILE_INFORMATION                 '//Pointer to FILE_INFORMATION struct
Public fData      As WIN32_FIND_DATA                  '//PTR to WIN32_FIND_DATA struct

Public Const VERSION = "2.0.0"                        '//Program Version (because I like doing it like this)
Public Const GUTMANN_PATTERN = 3                      '//Defines the Gutmann wipe method in our pattern array (frmOptions.cmdOK_Click)
Public Const TOTAL_PATTERNS = 4                       '//Total number of patterns we have in our arsenal see (frmOptions.cmdOK_Click)

'//Used for ShellExecute
Public Const SW_SHOWNORMAL = 1                        '//Activates and displays the window (whether minimized or maximized)

Public DOD_Pattern            As Integer              '//Overwrite Pattern (0,1,2)
Public Wipe_In_Progress       As Boolean              '//Are we currently wiping?
Public Rename_Files           As Boolean              '//Rename Files before delete?
Public debug_no_delete        As Boolean              '//Will not delete the files, just overwrite
Public blank_directory        As Boolean              '//Is there a empty directory structure loaded?


'//*
'// ----- Define Private Constants ------
'//*

Private Const GENERIC_READ = &H80000000               '//Specifies read access to the file. Data can be read from the file and the file pointer can be moved
Private Const GENERIC_WRITE = &H40000000              '//STANDARD_RIGHTS_WRITE
Private Const CREATE_ALWAYS = 2                       '//Creates a new file. The function overwrites the file if it exists.
Private Const OPEN_ALWAYS = 4                         '//Opens the file, if it exists. If the file does not exist, the function creates the file as if dwCreationDistribution were CREATE_NEW.
Private Const OPEN_EXISTING = 3                       '//Opens the file. The function fails if the file does not exist.
Private Const TRUNCATE_EXISTING = 5                   '//Opens the file. Once opened, the file is truncated so that its size is zero bytes. The calling process must open the file with at least GENERIC_WRITE access. The function fails if the file does not exist.
Private Const FILE_SHARE_WRITE = &H2                  '//Other open operations can be performed on the file for write access.
Private Const FILE_SHARE_READ = &H1                   '//Other open operations can be performed on the file for read access.
Private Const FILE_FLAG_WRITE_THROUGH = &H80000000    '//Instructs the operating system to write through any intermediate cache and go directly to the file. The operating system can still cache write operations, but cannot lazily flush them.
Private Const FILE_FLAG_NO_BUFFERING = &H20000000     '//request that its files not be cached

Private Const INVALID_HANDLE_VALUE = -1               '//-1 is invalid handle returned from CreateFile

'//File Attributes
Private Const FILE_ATTRIBUTE_READONLY = &H1           '//Read-Only file
Private Const FILE_ATTRIBUTE_HIDDEN = &H2             '//Hidden file
Private Const FILE_ATTRIBUTE_SYSTEM = &H4             '//The file is part of or is used exclusively by the operating system
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20           '//Archive file
Private Const FILE_ATTRIBUTE_NORMAL = &H80            '//Normal (no bits set)
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100        '//Temporary file
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800       '//The file or directory is compressed
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10         '//

'//Constants used for AlwaysOnTop() routine
Private Const SWP_NOMOVE = 2                          '//No move or no Resize flags
Private Const SWP_NOSIZE = 1                          '//No move or no Resize flags
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE        '//No move or no Resize flags
Private Const HWND_TOPMOST = -1                       '//Flag to make window ON TOP
Private Const HWND_NOTOPMOST = -2                     '//Flag to make window NOT ON TOP

'//*
'// ------ Windows API Declarations------
'//*

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileSpec As String) As Long
Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA)
Private Declare Function FindNextFile& Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA)
Private Declare Function FindClose& Lib "kernel32" (ByVal hFindFile As Long)

'//Private Variables
Private byteArray()       As Byte                     '//Array to store our overwrite data
Private SCbyteArray()     As Byte                     '//Array to store a single character to overwrite with
Private CMPbyteArray()    As Byte                     '//Second Array to Store Compliment values or byteArray()
Private RNDbyteArray()    As Byte                     '//Random byte array
Private GutmannArray1()   As Byte                     '//Re-Usable gutmann array 0x92 0x49 0x24
Private GutmannArray2()   As Byte                     '//Re-Usable gutmann array 0x49 0x24 0x92
Private GutmannArray3()   As Byte                     '//Re-Usable gutmann array 0x24 0x92 0x49
'
'

Public Function GutmannDelete(ByVal FILEname As String) As Long
  Dim i                 As Long             '//
  Dim pass              As Integer          '//
  Dim aByte             As Byte             '//
  Dim FILEsize          As Long             '//Size of the passed file (in bytes)
  Dim lBytesWritten     As Long             '//
  Dim array_ptr         As Long             '//
  Dim byte_ptr          As Integer          '//
  
  On Error GoTo Gutmann_Wipe_Error
  
  '//Make sure the files attributes are not read-only
  SetFileAttributes FILEname, FILE_ATTRIBUTE_NORMAL
  
  '//Get the size of the passed file
  FILEsize = FileLen(FILEname)
  
  '//Pass 1-4 Random bytes
  Do
    pass = pass + 1
    '//Fill random array with random bytes
    ReDim RNDbyteArray(FILEsize)
    For i = 0 To FILEsize
      aByte = rndint(0, 255)
      RNDbyteArray(i) = aByte
    Next i
    lBytesWritten = WriteArray(FILEname, RNDbyteArray())
    UpdateProgress pass, pass, 35
  Loop Until pass = 4
  
  '//Pass 5   0x55
  pass = pass + 1
  ReDim byteArray(FILEsize)
  For i = 0 To FILEsize
    byteArray(i) = &H55
  Next i
  lBytesWritten = WriteArray(FILEname, byteArray())
  UpdateProgress pass, pass, 35
  
  '//Pass 6   0xAA
  pass = pass + 1
  ReDim byteArray(FILEsize)
  For i = 0 To FILEsize
    byteArray(i) = &HAA
  Next i
  lBytesWritten = WriteArray(FILEname, byteArray())
  UpdateProgress pass, pass, 35
  
  '//Pass 7   0x92 0x49 0x24
  pass = pass + 1
  ReDim GutmannArray1(FILEsize)
  byte_ptr = 1
  Do
    Select Case byte_ptr
      Case 1: GutmannArray1(array_ptr) = &H92: byte_ptr = 2
      Case 2: GutmannArray1(array_ptr) = &H49: byte_ptr = 3
      Case 3: GutmannArray1(array_ptr) = &H24: byte_ptr = 1
    End Select
    array_ptr = array_ptr + 1
  Loop Until array_ptr = FILEsize
  lBytesWritten = WriteArray(FILEname, GutmannArray1())
  UpdateProgress pass, pass, 35
  
  '//Pass 8   0x49 0x24 0x92
  pass = pass + 1
  ReDim GutmannArray2(FILEsize)
  byte_ptr = 1
  array_ptr = 0
  Do
    Select Case byte_ptr
      Case 1: GutmannArray2(array_ptr) = &H49: byte_ptr = 2
      Case 2: GutmannArray2(array_ptr) = &H24: byte_ptr = 3
      Case 3: GutmannArray2(array_ptr) = &H92: byte_ptr = 1
    End Select
    array_ptr = array_ptr + 1
  Loop Until array_ptr = FILEsize
  lBytesWritten = WriteArray(FILEname, GutmannArray2())
  UpdateProgress pass, pass, 35


  '//Pass 9   0x24 0x92 0x49
  pass = pass + 1
  ReDim GutmannArray3(FILEsize)
  byte_ptr = 1
  array_ptr = 0
  Do
    Select Case byte_ptr
      Case 1: GutmannArray3(array_ptr) = &H24: byte_ptr = 2
      Case 2: GutmannArray3(array_ptr) = &H92: byte_ptr = 3
      Case 3: GutmannArray3(array_ptr) = &H49: byte_ptr = 1
    End Select
    array_ptr = array_ptr + 1
  Loop Until array_ptr >= FILEsize
  lBytesWritten = WriteArray(FILEname, GutmannArray3())
  UpdateProgress pass, pass, 35

  
  '//Pass 10 - 25  0x00 - 0xFF (step 0x11)
  pass = pass + 1
  aByte = &H0
  Do
    ReDim byteArray(FILEsize)
    For i = 0 To FILEsize
      byteArray(i) = aByte
    Next i
    lBytesWritten = WriteArray(FILEname, byteArray())
    UpdateProgress pass, pass, 35
    aByte = aByte + &H11
    pass = pass + 1
  Loop Until pass = 25
  
  '//Pass 26 0x92 0x49 0x24
  pass = pass + 1
  lBytesWritten = WriteArray(FILEname, GutmannArray1())
  UpdateProgress pass, pass, 35
  
  '//Pass 27 0x49 0x24 0x92
  pass = pass + 1
  lBytesWritten = WriteArray(FILEname, GutmannArray2())
  UpdateProgress pass, pass, 35
  
  '//Pass 28 0x24 0x92 0x49
  pass = pass + 1
  lBytesWritten = WriteArray(FILEname, GutmannArray3())
  UpdateProgress pass, pass, 35
  
  '//Pass 29 0x6D 0xB6 0xDB
  pass = pass + 1
  ReDim byteArray(FILEsize)
  byte_ptr = 1
  array_ptr = 0
  Do
    Select Case byte_ptr
      Case 1: byteArray(array_ptr) = &H6D: byte_ptr = 2
      Case 2: byteArray(array_ptr) = &HB6: byte_ptr = 3
      Case 3: byteArray(array_ptr) = &HDB: byte_ptr = 1
    End Select
    array_ptr = array_ptr + 1
  Loop Until array_ptr >= FILEsize
  lBytesWritten = WriteArray(FILEname, byteArray())
  UpdateProgress pass, pass, 35
  
  '//Pass 30 0xB6 0xDB 0x6D
  pass = pass + 1
  ReDim byteArray(FILEsize)
  byte_ptr = 1
  array_ptr = 0
  Do
    Select Case byte_ptr
      Case 1: byteArray(array_ptr) = &HB6: byte_ptr = 2
      Case 2: byteArray(array_ptr) = &HDB: byte_ptr = 3
      Case 3: byteArray(array_ptr) = &H6D: byte_ptr = 1
    End Select
    array_ptr = array_ptr + 1
  Loop Until array_ptr >= FILEsize
  lBytesWritten = WriteArray(FILEname, byteArray())
  UpdateProgress pass, pass, 35

  '//Pass 31 0xDB 0x6D 0xB6
  pass = pass + 1
  ReDim byteArray(FILEsize)
  byte_ptr = 1
  array_ptr = 0
  Do
    Select Case byte_ptr
      Case 1: byteArray(array_ptr) = &HDB: byte_ptr = 2
      Case 2: byteArray(array_ptr) = &H6D: byte_ptr = 3
      Case 3: byteArray(array_ptr) = &HB6: byte_ptr = 1
    End Select
    array_ptr = array_ptr + 1
  Loop Until array_ptr >= FILEsize
  lBytesWritten = WriteArray(FILEname, byteArray())
  UpdateProgress pass, pass, 35

  '//Pass 32-35 Random bytes
  Do
    pass = pass + 1
    '//Fill random array with random bytes
    ReDim RNDbyteArray(FILEsize)
    For i = 0 To FILEsize
      aByte = rndint(0, 255)
      RNDbyteArray(i) = aByte
    Next i
    lBytesWritten = WriteArray(FILEname, RNDbyteArray())
    UpdateProgress pass, pass, 35
  Loop Until pass = 35
  
  '//Return the number of bytes written (Using the last pass works fine)
  GutmannDelete = lBytesWritten
  
  Exit Function


Gutmann_Wipe_Error:
  Select Case Err.Number
    Case "6"    '//OVERFLOW
      '//The size of the file is probable to big (> 2147483647 bytes)
      ShowMessage MSG_FILE_TO_BIG, "GutmannDelete()"
      Exit Function
    Case "53"
      ShowMessage MSG_FILE_NOT_FOUND, "GutmannDelete()"
      Exit Function
    Case Else
      '//Display the error
      ShowMessage "Error: " & Err.Number & vbCrLf & "  Type: " & Err.Description, "GutmannDelete()"
  End Select

End Function

Public Function DoD_elete(ByVal FILEname As String, Pattern As Integer) As Long
  '// Standard DoD 5220.22-M:
  '//     Overwrite all addressable locations with a character,
  '//     its complement, then a random character and verify.
  
  On Error GoTo DOD_Wipe_Error
  
  Dim FILEsize          As Long             '//Size of the passed file (in bytes)
  Dim lBytesWritten     As Long             '//Returned from WriteFile API
  
  '//Make sure the files attributes are not read-only
  SetFileAttributes FILEname, FILE_ATTRIBUTE_NORMAL
  
  '//Get the size of the passed file
  FILEsize = FileLen(FILEname)
  
  '//We will first fill the byte arrays with the data we are going to overwrite.  This is
  '//where VB is the slowest.  So we load the arrays now, then we can just rotate through
  '//our overwrites.  LoadByteArrays loads the appropriate arrays for the overwrites.
  Select Case Pattern
    Case 0: UpdateProgress -1, 0, 11
    Case 1: UpdateProgress -1, 0, 6
    Case 2: UpdateProgress -1, 0, 2
  End Select
  
  LoadByteArrays FILEsize, Pattern

  '//Now that the arrays are filled, we can begin our writes.
  '//   c. Overwrite all addressable locations with a single character.
  '//   e. Overwrite all addressable locations with a character, its complement, then a random character.
  Select Case Pattern
    Case 0    '//US DoD 5220.22-M (8-306 /E, C and E)
      UpdateProgress 1, 4, 11
      lBytesWritten = WriteArray(FILEname, byteArray())
      UpdateProgress 2, 5, 11
      lBytesWritten = WriteArray(FILEname, CMPbyteArray())
      UpdateProgress 3, 6, 11
      lBytesWritten = WriteArray(FILEname, RNDbyteArray())
      UpdateProgress 4, 7, 11
      lBytesWritten = WriteArray(FILEname, SCbyteArray())
      UpdateProgress 5, 8, 11
      lBytesWritten = WriteArray(FILEname, byteArray())
      UpdateProgress 6, 9, 11
      lBytesWritten = WriteArray(FILEname, CMPbyteArray())
      UpdateProgress 7, 10, 11
      lBytesWritten = WriteArray(FILEname, RNDbyteArray())
      UpdateProgress 0, 11, 11
      
    Case 1    '//US DoD 5220.22-M (8-306 /E)
      UpdateProgress 1, 3, 6
      lBytesWritten = WriteArray(FILEname, byteArray())
      UpdateProgress 2, 4, 6
      lBytesWritten = WriteArray(FILEname, CMPbyteArray())
      UpdateProgress 3, 5, 6
      lBytesWritten = WriteArray(FILEname, RNDbyteArray())
      UpdateProgress 0, 6, 6
            
    Case 2    '//Pseudorandom Data
      UpdateProgress 1, 1, 2
      lBytesWritten = WriteArray(FILEname, RNDbyteArray())
      UpdateProgress 0, 2, 2
  End Select
  
  '//Return the number of bytes written (Using the last pass works fine)
  DoD_elete = lBytesWritten
  
  Exit Function


DOD_Wipe_Error:
  Select Case Err.Number
    Case "6"    '//OVERFLOW
      '//The size of the file is probable to big (> 2147483647 bytes)
      ShowMessage MSG_FILE_TO_BIG, "DoD_elete()"
      Exit Function
    Case "53"
      ShowMessage MSG_FILE_NOT_FOUND, "DoD_elete()"
      Exit Function
    Case Else
      '//Display the error
      ShowMessage "Error: " & Err.Number & vbCrLf & "  Type: " & Err.Description, "DOD_Elete()"
  End Select
End Function

Public Function rndint(ByVal lowerbound As Long, ByVal upperbound As Long) As Long
  '//Seed for random numbers
  Randomize Timer * 43
  
  rndint = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
  rndint = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Public Sub DeleteFile(aFile As String)
  On Error GoTo File_Delete_Error
  '//Make sure the files attributes are not read-only
  SetFileAttributes aFile, FILE_ATTRIBUTE_NORMAL
  
  '//Delete the file
  Kill aFile
  Exit Sub
  
File_Delete_Error:
  Select Case Err.Number
    Case 75: ShowMessage "Can not delete " & aFile & vbCrLf & "Make sure the file or folder is not in use.", "DeleteFile()"
    Resume Next
    Case Else: ShowMessage Err.Number & " -- " & Err.Description, "DeleteFile()"
  End Select
End Sub

Public Sub DeleteDirectory(aDirectory As String)
  Dim FSO, FS
  Set FSO = CreateObject("Scripting.FileSystemObject")
  FS = FSO.DeleteFolder(aDirectory, True)
End Sub

Private Function WriteArray(FILEname As String, anArray() As Byte) As Long
  '//This function uses the CreateFile and WriteFile API to write the
  '//array that was passed to the passed filename.  It returns the
  '//number of bytes written to the file.
  
  Dim fHandle         As Long
  Dim lBytesWritten   As Long
  Dim fSuccess        As Long
  Dim BytesToWrite    As Long
  
  '//Get file handle of file to overwrite
  'fHandle = CreateFile(FILEname, GENERIC_WRITE, FILE_SHARE_WRITE, 0, OPEN_ALWAYS, FILE_FLAG_WRITE_THROUGH, 0)
  fHandle = CreateFile(FILEname, GENERIC_WRITE, 0, 0, OPEN_ALWAYS, FILE_FLAG_WRITE_THROUGH, 0)
  If fHandle = INVALID_HANDLE_VALUE Then
    '//Invalid File Handle
    CloseHandle fHandle
  Else
    '//Get the length of data to write
    BytesToWrite = (UBound(anArray) * LenB(anArray(0)))
    
    '//Overwrite the file with the byte array
    fSuccess = WriteFile(fHandle, anArray(LBound(anArray)), BytesToWrite, lBytesWritten, 0)
    
    '//Flush any buffers the system used for the file
    FlushFileBuffers fHandle
    CloseHandle fHandle
  End If
  
  DoEvents
  WriteArray = lBytesWritten
End Function

Public Function CRC16Cipher(Text As String, Mask As String)
  
  '// This function generates an 8 byte "Ciphered" code (determined by Mask)
  '// for the passed Text.  The passed mask is actually filled in with CRC value
  '// and gets returned from this routine.
  '// Mask options are:
  '//   "$"   -   Letter
  '//   "#"   -   Number
  '//   "?"   -   Random
  
  Dim CS              As String
  Dim CRCText         As String
  Dim TempChar        As String
  Dim Power(0 To 7)   As Integer
  Dim CRC             As Long
  Dim ByteVal         As Long
  Dim TestBit         As Integer
  Dim LC              As Integer
  Dim CharPos         As Integer
  Dim i, j            As Integer
      
  Const NUMBER_SET = "1234567890"
  Const CHAR_SET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
  '//Generate CRC Look-up-Table
  For i = 0 To 7: Power(i) = 2 ^ i: Next i
  
  CRCText = Text + Mask
  For i = 1 To Len(CRCText)
    ByteVal = Asc(Mid$(CRCText, i, 1))
    For j = 7 To 0 Step -1
      TestBit = ((CRC And 32768) = 32768) Xor ((ByteVal And Power(j)) = Power(j))
      CRC = ((CRC And 32767&) * 2&)
      If TestBit Then CRC = CRC Xor &H8005&
    Next j
  Next i
  CRCText = ""
  
  TempChar = LTrim$(RTrim$(Str$(CRC)))
  Dim Map(4): For i = 0 To 4: Map(i) = Val(Mid$(TempChar, i + 1, 1)): Next i
    
  For i = 1 To Len(Mask)
    Select Case Mid$(Mask, i, 1)
      Case "?": CS = NUMBER_SET + CHAR_SET
      Case "#": CS = NUMBER_SET
      Case "$": CS = CHAR_SET
      Case Else: CS = " "
    End Select

    If CS <> " " Then
      LC = Len(CS)
      CharPos = ((Map(i Mod 5) * Asc(Mid$(CS, (i Mod LC) + 1, 1)) + 1))
      Mid$(Mask, i, 1) = Mid$(CS, (CharPos Mod LC) + 1, 1)
    End If
  Next i
  CRC16Cipher = Mask
End Function

Public Function RenameFile(FILEname As String, NumberOfTimes As Integer) As String
  Dim i             As Integer
  Dim Encode        As String
  Dim NewFileName   As String
  Dim LastFileName  As String
  Dim pos           As Integer
  Dim FILEpath      As String
  
  On Error GoTo Rename_File_Error
  
  '//Save initial filename so we know what we are renaming
  LastFileName = FILEname
  
  '//Get path from file name
  pos = InStrRev(FILEname, "\")
  FILEpath = Mid$(FILEname, 1, pos)
  
  For i = 1 To NumberOfTimes
    Encode = FILEname & Trim$(Str$(Timer))
    NewFileName = CRC16Cipher(Encode, "$#$#$$#$")
    
    '//Rename the file
    Name LastFileName As FILEpath & NewFileName
    
    '//Save the new filename so we know what file we are renaming next pass
    LastFileName = FILEpath & NewFileName
  Next i

  '//Return the last file name created
  RenameFile = FILEpath & NewFileName
  Exit Function

Rename_File_Error:
  Select Case Err.Number
    Case 75: ShowMessage "Can not rename " & FILEname & vbCrLf & "Make sure the file is not in use.", "RenameFile()"
    Resume Next
    Case Else: ShowMessage Err.Number & " -- " & Err.Description, "RenameFile()"
  End Select
End Function

Public Sub GetFileInformation(FILEname As String)
  '// This sub sets up the public type FILE_INFO.  I got the idea
  '// for this code from Microsoft Knowledge Base Article ID: Q154821
  
  '// This sub gets called to get information about a single file
  
  Dim fHandle     As Long
  Dim f_creation  As FILETIME
  Dim f_access    As FILETIME
  Dim f_modify    As FILETIME
  Dim sys_time    As SYSTEMTIME
  Dim ret         As Long
  
  fHandle = CreateFile(FILEname, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
  
  If fHandle = INVALID_HANDLE_VALUE Then
    '//Invalid File Handle
    CloseHandle fHandle
  Else
    '//Set the file name in struct
    FILE_INFO.fName = FILEname
    
    '//Get the size of the passed file
    FILE_INFO.fSize = FileLen(FILEname)
  
    '//Retrieve the date and time that the file was created, last accessed, and last modified.
    ret = GetFileTime(fHandle, f_creation, f_access, f_modify)
  
    '//Convert file times based on the Coordinated Universal Time (UTC) to local file time
    ret = FileTimeToLocalFileTime(f_creation, f_creation)
    ret = FileTimeToLocalFileTime(f_access, f_access)
    ret = FileTimeToLocalFileTime(f_modify, f_modify)
  
    '//Convert 64-bit file times to system time format
    ret = FileTimeToSystemTime(f_creation, sys_time)
    FILE_INFO.CreationDate = CDate(sys_time.wMonth & "/" & sys_time.wDay & "/" & sys_time.wYear)
    ret = FileTimeToSystemTime(f_access, sys_time)
    FILE_INFO.AccessDate = CDate(sys_time.wMonth & "/" & sys_time.wDay & "/" & sys_time.wYear)
    ret = FileTimeToSystemTime(f_modify, sys_time)
    FILE_INFO.ModifyDate = CDate(sys_time.wMonth & "/" & sys_time.wDay & "/" & sys_time.wYear)
    
    '//Finished getting info from the file.  Close it up!
    CloseHandle fHandle
  End If
End Sub

Public Function IsDirectory(aFilePath As String) As Boolean
  '//This function determines if the passed file path is a directory or not
  
  Dim fAttribs As Long
  
  fAttribs = GetFileAttributes(aFilePath)
  If fAttribs = FILE_ATTRIBUTE_DIRECTORY Then
    IsDirectory = True
  Else
    IsDirectory = False
  End If
End Function

Function FileExists(FILEname) As Boolean
  '//Determine if a file exists or not
  
  On Error Resume Next
  
  FileExists = False
  FileExists = (Dir(FILEname) <> "")
End Function

Public Sub GetAllFilesInDir(ByVal sFolderPath As String)
  '/////// note: to stop this use bCancelFileListAction = True

  Dim fHandle       As Long                 '//File Handle PTR
  Dim ret           As Long                 '//Return value of FindNextFile
  Dim FILEname      As String               '//Stores File path and name

  '//Get file handle (returns null-terminated string)
  fHandle = FindFirstFile(sFolderPath, fData)
  
  '//Exit if there was an Error Getting First Entry
  If fHandle = INVALID_HANDLE_VALUE Then
    '//Close the file
    FindClose (fHandle)
    Exit Sub
  End If
  
  '//Initialize the return value for the FindNextFile function
  ret = 1
  
  Do While ret <> 0
    '//Remove the null characters from Returned FILEname
    FILEname = StripNulls(fData.cFileName)
  
    '//If it is a Directory but NOT the "." or ".." directories
    If FILEname <> "." And FILEname <> ".." And fData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
      '//See whats inside this directory ...
      GetAllFilesInDir Mid$(sFolderPath, 1, Len(sFolderPath) - 3) & FILEname & "\*.*"
    Else        'If the item is not a directory (folder) and...
      If FILEname <> "." And FILEname <> ".." And FILEname <> "" And Not fData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
        '//OK - So we have a file, now we will load the FILE_INFO struct with the file info!
        
        '//Set the file path & name
        FILE_INFO.fName = Mid$(sFolderPath, 1, Len(sFolderPath) - 3) & FILEname
      
        '//Setup the FILE_INFO struct
        GetFileInformation FILE_INFO.fName
        
        '//Display Items in the listview control
        AddFileToListView FILE_INFO
        
        DoEvents
      End If
    End If
      
    '//Get Next Entry
    ret = FindNextFile(fHandle, fData)
    
    If ret = 0 Then
      '//No more files, Close Handle
      FindClose (fHandle)
      Exit Sub
    End If
            
    DoEvents
  Loop

  '//Close Handle
  FindClose (fHandle)
End Sub

Private Function StripNulls(ByVal FileWithNulls As String) As String
  Dim NullPos As Integer
  
  NullPos = InStr(1, FileWithNulls, vbNullChar, 0)
  
  If NullPos <> 0 Then
    StripNulls = Left(FileWithNulls, NullPos - 1)
  End If
End Function

Public Function Compliment(ByVal DecimalNum As Byte) As Byte
  Dim i             As Integer
  Dim bitbuffer     As String
  Dim len_test      As Integer
  Dim test_bit      As String
  Dim cmp_buffer    As String
  Dim bits          As Long
  
  '//Remove and spaces in the bitbuffer
  bitbuffer = Trim$(Str$(DecimalNum Mod 2))
  DecimalNum = DecimalNum \ 2

  Do While DecimalNum <> 0
    bitbuffer = Trim$(Str$(DecimalNum Mod 2)) & bitbuffer
    DecimalNum = DecimalNum \ 2
  Loop

  '//Now see if we need to adjust the length to make 8 bits
  len_test = 8 - Len(bitbuffer)
  Select Case len_test
    Case 1: bitbuffer = "0" & bitbuffer
    Case 2: bitbuffer = "00" & bitbuffer
    Case 3: bitbuffer = "000" & bitbuffer
    Case 4: bitbuffer = "0000" & bitbuffer
    Case 5: bitbuffer = "00000" & bitbuffer
    Case 6: bitbuffer = "000000" & bitbuffer
    Case 7: bitbuffer = "0000000" & bitbuffer
  End Select
  
  '//At this point, the bitbuffer contains a string representing our
  '//passed decimal number.  Now we will compliment the byte!@!

  For i = 1 To Len(bitbuffer)
    test_bit = Mid$(bitbuffer, i, 1)
    Select Case test_bit
      Case "0": cmp_buffer = cmp_buffer & "1"
      Case "1": cmp_buffer = cmp_buffer & "0"
    End Select
  Next i

  '//Now convert the complimented binary string back to a value
  For i = 1 To Len(cmp_buffer)
    bits = bits + (Mid$(cmp_buffer, Len(cmp_buffer) - i + 1, 1) * (2 ^ (i - 1)))
  Next i

  '//Return the complimented value
  Compliment = bits
End Function

Public Sub AddFileToListView(ByRef fInfo As FILE_INFORMATION)
  Dim objLvi As MSComctlLib.ListItem
  
  '//Display Items in the listview control
  Set objLvi = frmMain.lvFiles.ListItems.Add()
  objLvi.Text = fInfo.fName
  objLvi.SubItems(1) = fInfo.fSize
  objLvi.SubItems(2) = fInfo.AccessDate
  objLvi.SubItems(3) = fInfo.CreationDate
  objLvi.SubItems(4) = fInfo.ModifyDate
  
  Set objLvi = Nothing

  frmMain.sb.Panels.Item(2).Text = "Files: " & Trim$(Str$(frmMain.lvFiles.ListItems.Count))
  frmMain.sb.Refresh
End Sub

Private Sub LoadByteArrays(sizeofArray As Long, Pattern As Integer)
  '// c. Overwrite all addressable locations with a single character.
  '// e. Overwrite all addressable locations with a character, its complement, then a random character.
  
  Dim aByte As Byte
  Dim CMPbyte As Byte
  Dim RNDbyte As Byte
  Dim SCbyte As Byte
  Dim i As Long
  
  Select Case Pattern
    Case 0    '//US DoD 5220.22-M (8-306 /E, C and E)
      '//Get a random byte
      aByte = rndint(0, 255)
      
      '//Get the compliment of that byte
      CMPbyte = Compliment(aByte)
      
      '//Fill byte array with data
      ReDim byteArray(sizeofArray)
      For i = 0 To sizeofArray
        byteArray(i) = aByte
      Next i
        
      UpdateProgress -1, 1, 11
      
      '//Fill Compliment byte array
      ReDim CMPbyteArray(sizeofArray)
      For i = 0 To sizeofArray
        CMPbyteArray(i) = CMPbyte
      Next i
      
      UpdateProgress -1, 2, 11
      
      '//Now we will fill random array with random bytes
      ReDim RNDbyteArray(sizeofArray)
      For i = 0 To sizeofArray
        RNDbyte = rndint(0, 255)
        RNDbyteArray(i) = RNDbyte
      Next i
      
      UpdateProgress -1, 3, 11
      
      '//Fill Single Character array with a char
      '//Get a character
      SCbyte = rndint(0, 255)
      ReDim SCbyteArray(sizeofArray)
      For i = 0 To sizeofArray
        SCbyteArray(i) = SCbyte
      Next i
      
      UpdateProgress -1, 4, 11
      
    Case 1    '//US DoD 5220.22-M (8-306 /E)
      '//Get a random byte
      aByte = rndint(0, 255)
      
      '//Get the compliment of that byte
      CMPbyte = Compliment(aByte)
      
      '//Fill byte array with data
      ReDim byteArray(sizeofArray)
      For i = 0 To sizeofArray
        byteArray(i) = aByte
      Next i
      
      UpdateProgress -1, 1, 6
      
      '//Fill Compliment byte array
      ReDim CMPbyteArray(sizeofArray)
      For i = 0 To sizeofArray
        CMPbyteArray(i) = CMPbyte
      Next i
      
      UpdateProgress -1, 2, 6
      
      '//Now we will fill random array with random bytes
      ReDim RNDbyteArray(sizeofArray)
      For i = 0 To sizeofArray
        RNDbyte = rndint(0, 255)
        RNDbyteArray(i) = RNDbyte
      Next i
      
      UpdateProgress -1, 3, 6
    
    Case 2    '//Pseudorandom Data
      '//Now we will fill random array with random bytes
      ReDim RNDbyteArray(sizeofArray)
      For i = 0 To sizeofArray
        RNDbyte = rndint(0, 255)
        RNDbyteArray(i) = RNDbyte
      Next i
      
      UpdateProgress -1, 1, 2
  End Select
End Sub

Public Function TruncateFilename(FILEpath As String, MaxLength As Integer) As String
  If (Len(FILEpath) > MaxLength) Then
    Dim final As String
    Dim pos As Integer
    
    pos = InStrRev(FILEpath, "\")
    pos = pos - 1
    
    pos = InStrRev(FILEpath, "\", pos)
    
    final = Left$(FILEpath, 3) & "..." & Mid$(FILEpath, pos)
    TruncateFilename = final
  Else
    TruncateFilename = FILEpath
  End If
End Function

Private Sub UpdateProgress(pass As Integer, Stage As Integer, TotalStages As Integer)
  With frmWiping
    Select Case pass
      Case -1:    .lblPass.Caption = "Loading data arrays..."
      Case 0:     .lblPass.Caption = ""
      Case Else:  .lblPass.Caption = pass
    End Select
    
    .lblFilePercent.Caption = Int((Stage / TotalStages) * 100) & " %"
    .pbFile.value = (Stage / TotalStages) * 100
    
    .Refresh
  End With
  DoEvents
End Sub

Public Sub AlwaysOnTop(FrmID As Form, OnTop As Boolean)
  Select Case OnTop
    Case True:    OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Case False:   OnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
  End Select
End Sub

Public Sub ShowMessage(Message As String, Optional Caption As String)
  Load frmMessage
  
  With frmMessage
    If Caption <> "" Then .Caption = Caption
    .lblMessage.Caption = Message
    .Show
  End With
  
  '//Make the message top most
  AlwaysOnTop frmMessage, True
End Sub

Public Sub UpdateStatusBarBytePattern()
  Select Case DOD_Pattern
    Case 0: frmMain.sb.Panels(1).Text = " Byte Pattern:  --  " & MSG_PATTERN_0
    Case 1: frmMain.sb.Panels(1).Text = " Byte Pattern:  --  " & MSG_PATTERN_1
    Case 2: frmMain.sb.Panels(1).Text = " Byte Pattern:  --  " & MSG_PATTERN_2
    Case 3: frmMain.sb.Panels(1).Text = " Byte Pattern:  --  " & MSG_PATTERN_3
  End Select
End Sub
