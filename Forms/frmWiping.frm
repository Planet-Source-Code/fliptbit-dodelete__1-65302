VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWiping 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DODelete"
   ClientHeight    =   3912
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   4548
   ControlBox      =   0   'False
   Icon            =   "frmWiping.frx":0000
   LinkTopic       =   "frmWiping"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3912
   ScaleWidth      =   4548
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbTotal 
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   3612
      _ExtentX        =   6371
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pbFile 
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3612
      _ExtentX        =   6371
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdStopWipe 
      Caption         =   "Stop"
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   3360
      Width           =   1452
   End
   Begin VB.Label lblTotalPercent 
      Alignment       =   2  'Center
      Caption         =   "100 %"
      Height          =   252
      Left            =   3840
      TabIndex        =   11
      Top             =   2880
      Width           =   492
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   612
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblFilePercent 
      Alignment       =   2  'Center
      Caption         =   "100 %"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   1800
      Width           =   492
   End
   Begin VB.Label lblPass 
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
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   3252
   End
   Begin VB.Label lblPasslbl 
      Caption         =   "Pass:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   492
   End
   Begin VB.Label lblFileName 
      Height          =   252
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   3132
   End
   Begin VB.Label lblItemlbl 
      Caption         =   "File:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   492
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblWipePattern 
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
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   3132
   End
   Begin VB.Label lblWipeTypelbl 
      Caption         =   "Wipe Pattern:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "frmWiping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//-----------------------------------------------------------------------------------------------------------------
'//    Project: DODelete.vbp
'//   FileName: frmMain.frm
'//   Coded By: FliptBit
'//
'// Start Date: 26 MAY 2001
'//
'//
'//    Purpose: This form simply displays the current wipe operation and displays the overall
'//             progress of the wiping.  Also allows the user to stop wiping the files.
'//
'//
'//     To Do/Issues:
'//
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

Private Sub cmdStopWipe_Click()
  '//Reset the wipe in progress flag
  Wipe_In_Progress = False
  
  cmdStopWipe.Enabled = False
  cmdStopWipe.Caption = "Stopping..."
  cmdStopWipe.Refresh
  Me.Refresh
  DoEvents
End Sub

Private Sub Form_Load()
  '//Setup the progress bars min and max values
  pbFile.Min = 0:   pbFile.Max = 100
  pbTotal.Min = 0:  pbTotal.Max = 100

  '//Set % labels to initial value
  lblFilePercent.Caption = "0 %"
  lblTotalPercent.Caption = "0 %"

  '//Tell user the pattern we are using
  With lblWipePattern
    Select Case DOD_Pattern
      Case 0: .Caption = MSG_PATTERN_0
      Case 1: .Caption = MSG_PATTERN_1
      Case 2: .Caption = MSG_PATTERN_2
      Case 3: .Caption = MSG_PATTERN_3
    End Select
  End With
End Sub

