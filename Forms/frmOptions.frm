VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DODelete Options"
   ClientHeight    =   3276
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   5340
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3276
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3960
      TabIndex        =   6
      Top             =   2760
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   2880
      TabIndex        =   5
      Top             =   2760
      Width           =   972
   End
   Begin VB.CheckBox chkRenameFiles 
      Caption         =   "Rename files before delete"
      Height          =   252
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   2412
   End
   Begin VB.Frame fraPattern 
      Caption         =   "Select Overwrite Pattern"
      Height          =   1932
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4932
      Begin VB.CheckBox chkPattern 
         Caption         =   "Gutmann Wipe"
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   2292
      End
      Begin VB.CheckBox chkPattern 
         Caption         =   "Pseudorandom Data"
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   2292
      End
      Begin VB.CheckBox chkPattern 
         Caption         =   "US DoD 5220.22-M (8-306 /E)"
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   3012
      End
      Begin VB.CheckBox chkPattern 
         Caption         =   "US DoD 5220.22-M (8-306 /E, C and E)"
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   3252
      End
      Begin VB.Label lblPasses 
         Caption         =   "35 Passes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   852
      End
      Begin VB.Label lblPasses 
         Caption         =   "  1 Pass"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Top             =   1440
         Width           =   852
      End
      Begin VB.Label lblPasses 
         Caption         =   "  3 Passes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   3840
         TabIndex        =   9
         Top             =   1080
         Width           =   852
      End
      Begin VB.Label lblPasses 
         Caption         =   "  7 Passes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   3840
         TabIndex        =   8
         Top             =   720
         Width           =   852
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//-----------------------------------------------------------------------------------------------------------------
'//    Project: DODelete.vbp
'//   FileName: frmOptions.frm
'//   Coded By: FliptBit
'//
'// Start Date: 18 MAY 2006
'//
'//
'//    Purpose: This form is used to setup the overwrite patterns, and other settings
'//             to be used in this program.
'//
'//
'//     To Do/Issues:
'//       - Add code to save/restore settings
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
'
'

Private Sub chkPattern_Click(Index As Integer)
  Dim check As CheckBox
  If chkPattern(Index).value = 1 Then
    For Each check In chkPattern
      If check.Index <> Index Then check.value = 0
    Next
  End If
End Sub

Private Sub cmdCancel_Click()
  frmMain.Show
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim i As Integer
  
  For i = 0 To TOTAL_PATTERNS - 1
    If chkPattern(i).value = 1 Then
      DOD_Pattern = i
      RegWriteOverwritePattern Trim$(Str$(i))
    End If
  Next i
  
  If chkRenameFiles.value = 1 Then
    Rename_Files = True
    RegWriteRenameFiles "1"
  Else
    Rename_Files = False
    RegWriteRenameFiles "0"
  End If
  
  UpdateStatusBarBytePattern
  
  Unload Me
  frmMain.Show
End Sub

Private Sub Form_Load()
  '//Display current settings
  chkPattern(DOD_Pattern) = 1
  
  If Rename_Files = True Then
    chkRenameFiles.value = 1
  Else
    chkRenameFiles.value = 0
  End If
End Sub

