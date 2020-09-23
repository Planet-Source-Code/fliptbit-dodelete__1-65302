Attribute VB_Name = "modMessages"
'//-----------------------------------------------------------------------------------------------------------------
'//    Project: DODelete.vbp
'//   FileName: modMessages.bas
'//   Coded By: FliptBit
'//
'// Start Date: 17 MAY 2006
'//
'//
'//    Purpose: This module will contain most messages displayed to the user.  This
'//             helps to consolidate all the messages in one place for easy replacement,
'//             improvements, etc.
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

'//Error messages
Public Const MSG_FILE_TO_BIG = "The file you are trying to wipe is too big for this program."
Public Const MSG_FILE_NOT_FOUND = "The file you are trying to wipe can not be found."

'//Misc program messages
Public Const MSG_NO_FILES_TO_WIPE = "Please drag any files/folders you want to wipe into the file list."
Public Const MSG_STARTUP_HELP = "Drag files you want to wipe here..."
Public Const MSG_DELETE_ALL_FILES = "Are you sure you want to wipe these files??   Once these files are deleted they can not be recovered!"
Public Const MSG_DELETE_THIS_FILE = "Are you sure you want to wipe this file??   Once this file is deleted it can not be recovered!"
Public Const MSG_NO_VALILD_FILES = "The directory does not contain any files."

Public Const MSG_APP_DESCRIPTION = "This program will completely wipe files making them impossible to recover even with professional software recovery programs.  BE CAREFUL!!!"
Public Const MSG_APP_COPYRIGHT = "© 2006 FliptBit Technologies, Inc."

Public Const MSG_PATTERN_0 = "US DoD 5220.22-M (8-306 /E, C and E)"
Public Const MSG_PATTERN_1 = "US DoD 5220.22-M (8-306 /E)"
Public Const MSG_PATTERN_2 = "Pseudorandom Data"
Public Const MSG_PATTERN_3 = "Peter Gutmann's Overwrite Method"


