VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DODelete"
   ClientHeight    =   1704
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6612
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1704
   ScaleWidth      =   6612
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   1452
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   672
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6372
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//-----------------------------------------------------------------------------------------------------------------
'//    Project: DODelete.vbp
'//   FileName: frmMessages.frm
'//   Coded By: FliptBit
'//
'// Start Date: 19 MAY 2006
'//
'//
'//    Purpose: This form will display a message to the user.  It is used instead of
'//             the typical VB MsgBox function.  I hate that function.  Try writing a
'//             serial communications program and use the MsgBox function.  See what
'//             happens when the message box is displayed with MsgBox --- Your program
'//             stops processing all code - even timer and S0 interrupts and such.
'//
'//             I'll stop complaining now...
'//
'//       Note: This form is loaded from ShowMessage(Message As String, Optional Caption As String)
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

Private Sub cmdOK_Click()
  frmMain.Show
  Unload Me
End Sub
