VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3948
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   5544
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2718.098
   ScaleMode       =   0  'User
   ScaleWidth      =   5200.658
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   384
      Left            =   240
      Picture         =   "frmAbout.frx":0742
      ScaleHeight     =   263.118
      ScaleMode       =   0  'User
      ScaleWidth      =   263.118
      TabIndex        =   1
      Top             =   240
      Width           =   384
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4128
      TabIndex        =   0
      Top             =   2988
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4140
      TabIndex        =   2
      Top             =   3432
      Width           =   1245
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.fliptbit.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   192
      Left            =   600
      MouseIcon       =   "frmAbout.frx":100C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3000
      Width           =   1152
   End
   Begin VB.Label lblMailto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "webmaster@fliptbit.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   192
      Left            =   600
      MouseIcon       =   "frmAbout.frx":115E
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3240
      Width           =   1716
   End
   Begin VB.Label lblHttplbl 
      Caption         =   "  http:"
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   492
   End
   Begin VB.Label lblMaillbl 
      Caption         =   "email:"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   492
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   744
      Left            =   960
      MouseIcon       =   "frmAbout.frx":12B0
      Picture         =   "frmAbout.frx":1F7A
      Top             =   120
      Width           =   2556
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -22.514
      X2              =   5202.534
      Y1              =   1933.237
      Y2              =   1933.237
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   816
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   3888
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -11.257
      X2              =   5199.72
      Y1              =   1949.76
      Y2              =   1949.76
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   960
      TabIndex        =   5
      Top             =   840
      Width           =   2556
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "© 2006 FliptBit Technologies, Inc."
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   2400
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//-----------------------------------------------------------------------------------------------------------------
'//    Project: DODelete.vbp
'//   FileName: frmAbout.frm
'//   Coded By: Microsoft
'//
'// Start Date: 23 MAY 2006
'//
'//
'//    Purpose: Display information about the program and system
'//
'//      Notes: Typical VB self-generated About form, with my own additions
'//
'//
'//******************************************************************************************************************
'//
'//    License:
'//
'//    Copyright (c) XXXX  Microsoft Corporation  ( I guess )
'//
'//-----------------------------------------------------------------------------------------------------------------


Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'
'

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = "About " & App.Title
  lblVersion.Caption = "Version " & VERSION
  
  lblDisclaimer.Caption = MSG_APP_COPYRIGHT
  lblDescription.Caption = MSG_APP_DESCRIPTION
End Sub

Private Sub lblMailto_Click()
  Dim ret As Long
  ret = ShellExecute(Me.hwnd, vbNullString, "mailto:support@fliptbit.com", vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Private Sub lblURL_Click()
  Dim ret As Long
  ret = ShellExecute(Me.hwnd, vbNullString, "http://www.fliptbit.com", vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Public Sub StartSysInfo()
  On Error GoTo SysInfoErr

  Dim rc As Long
  Dim SysInfoPath As String
  
  ' Try To Get System Info Program Path\Name From Registry...
  If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
  ' Try To Get System Info Program Path Only From Registry...
  ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    ' Validate Existance Of Known 32 Bit File Version
    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        
    ' Error - File Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
  ' Error - Registry Entry Can Not Be Found...
  Else
      GoTo SysInfoErr
  End If
  
  Call Shell(SysInfoPath, vbNormalFocus)
  
  Exit Sub
SysInfoErr:
  ShowMessage "System Information Is Unavailable At This Time", App.Title
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

