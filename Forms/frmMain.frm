VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DODelete"
   ClientHeight    =   4332
   ClientLeft      =   120
   ClientTop       =   804
   ClientWidth     =   11292
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4332
   ScaleWidth      =   11292
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   3852
      Left            =   120
      ScaleHeight     =   3804
      ScaleWidth      =   924
      TabIndex        =   5
      Top             =   120
      Width           =   972
      Begin VB.Label lblWipe 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wipe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   732
      End
      Begin VB.Image imgControl 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   408
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmMain.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1594
         ToolTipText     =   "Securely wipe all files in the list"
         Top             =   360
         Width           =   408
      End
      Begin VB.Label lblClearList 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   732
      End
      Begin VB.Image imgControl 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   408
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmMain.frx":25D6
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":32A0
         ToolTipText     =   "Clear all files from the list"
         Top             =   1200
         Width           =   408
      End
      Begin VB.Image imgControl 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   408
         Index           =   2
         Left            =   240
         MouseIcon       =   "frmMain.frx":42E2
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":4FAC
         ToolTipText     =   "Display program settings"
         Top             =   2280
         Width           =   408
      End
      Begin VB.Image imgControl 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   408
         Index           =   3
         Left            =   240
         MouseIcon       =   "frmMain.frx":5FEE
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":6CB8
         ToolTipText     =   "Display help"
         Top             =   3120
         Width           =   408
      End
      Begin VB.Label lblSettings 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   732
      End
      Begin VB.Label lblHelp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   732
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   960
         Y1              =   1800
         Y2              =   1800
      End
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   324
      Left            =   0
      TabIndex        =   2
      Top             =   4008
      Width           =   11292
      _ExtentX        =   19918
      _ExtentY        =   572
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17144
            MinWidth        =   17144
            Picture         =   "frmMain.frx":7CFA
            Text            =   "DODelete Version 2.0.0"
            TextSave        =   "DODelete Version 2.0.0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Files:"
            TextSave        =   "Files:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraFiles 
      Height          =   3972
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   9972
      Begin VB.FileListBox File1 
         Height          =   1032
         Left            =   8400
         TabIndex        =   4
         Top             =   1920
         Width           =   1092
      End
      Begin VB.ListBox lstDirectories 
         Height          =   1200
         Left            =   8400
         TabIndex        =   3
         Top             =   600
         Width           =   1092
      End
      Begin MSComctlLib.ListView lvFiles 
         Height          =   3612
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9732
         _ExtentX        =   17166
         _ExtentY        =   6371
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   9596
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size (bytes)"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Accessed"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Created"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Modified"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help &Topics"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About D0Delete..."
      End
   End
   Begin VB.Menu mnulvPopUp 
      Caption         =   "lvPopup"
      Visible         =   0   'False
      Begin VB.Menu mnulvWipeAll 
         Caption         =   "Wipe All Files"
      End
      Begin VB.Menu mnulvWipeSelected 
         Caption         =   "Wipe Selected"
      End
      Begin VB.Menu mnulvSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnulvRemoveItem 
         Caption         =   "Remove from list"
      End
      Begin VB.Menu mnulvRemoveAll 
         Caption         =   "Remove All"
      End
   End
End
Attribute VB_Name = "frmMain"
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
'//    Purpose: This is my attempt to adhere to the DOD 5220.22-M standard as stated
'//             in the Nation Industrial Security Program Operating Manual (REV Jan 1995).
'//             See material in Related Documents ---------------->
'//
'//
'//     To Do/Issues:
'//       - Add code for more command line options
'//       - Add code for more/better error handling
'//       - Fix lvFiles to accept more than one drag-n-dop occurence.  As of now, for every drag-n-drop
'//         the list is cleared and regenerated with the newly dropped data.
'//       - Add code to do cool stuff with the ListView control (sorting, up/down arrows, etc.)
'//
'//
'//      Notes: This program has drag-n-drop capabilites.
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

Private Sub Form_Load()
  frmMain.Caption = "DODelete -- Version" & VERSION
    
  '//Hide file & directory array controls
  lstDirectories.Visible = False
  File1.Visible = False
  
  '//Load user settings from registry
  LoadRegistrySettings
  
  '//Set defaults
  debug_no_delete = False
  
  UpdateStatusBarBytePattern

  AddHelpMSGtoList
End Sub

Private Sub imgControl_Click(Index As Integer)
  '//If we are already wiping than ignore the click
  If Wipe_In_Progress = True Then Exit Sub
  
  Select Case Index
    Case 0
      '//Wipe Files
      
      Dim i               As Long               '//Standard INC
      Dim FILEname        As String             '//The File Name
      Dim StartTime       As Single             '//Used to time the routine
      Dim EndTime         As Single             '//Used to tim ethe routine
      Dim lBytesWritten   As Long               '//Number of bytes written to file
      Dim TotalFiles      As Long               '//Total number of files to wipe
      Dim TotalBytes      As Long               '//Total number of bytes written
      Dim FileCount       As Long               '//Counter for progress bar
      Dim ff              As Integer            '//Free file handle (for making file lengths 0 bytes)
      
      '//See if there is any files to delete.  If not tell the user.
      If lvFiles.ListItems.Item(1) = MSG_STARTUP_HELP Then
        '//User has not selected a file to wipe, so tell them!!
        ShowMessage MSG_NO_FILES_TO_WIPE, "File Not Found"
        Exit Sub
      End If
      
      If MsgBox(MSG_DELETE_ALL_FILES, vbYesNo + vbCritical, "Wipe Files") = vbNo Then
        '//User did not want to commit to the wipe
        Exit Sub
      End If
      
      '//Set flag that we are wiping
      Wipe_In_Progress = True
      
      frmMain.MousePointer = vbHourglass
      DoEvents
    
      '//Set the start time for benchmarking
      StartTime! = Timer
      
      '//Get total number of files to wipe (for progress bar calculations)
      TotalFiles = lvFiles.ListItems.Count
      
      '//Load the wiping progress form
      Load frmWiping: frmWiping.Show
      DoEvents
      
      Do While lvFiles.ListItems.Count > 0 And Wipe_In_Progress = True
        '//Grab the first line from the listview (could use last too)
        lvFiles.ListItems.Item(1).Selected = True
        lvFiles.Refresh
        FILEname = lvFiles.ListItems.Item(1)
        
        '//Truncate the name if needed so it shows on our screen
        frmWiping.lblFileName.Caption = TruncateFilename(FILEname, 25)
        
        '// ***** Overwrite the file *****
        If DOD_Pattern <> GUTMANN_PATTERN Then
          '//A DOD pattern is being used
          lBytesWritten = DoD_elete(FILEname, DOD_Pattern)
        Else
          '//The Gutmann patterns are being used
          lBytesWritten = GutmannDelete(FILEname)
        End If
        
        DoEvents
        
        '//Increment our filecounter and # of bytes (for percentages)
        FileCount = FileCount + 1
        TotalBytes = TotalBytes + lBytesWritten
        
        '//Now we will make the length of the overwritten file 0 bytes
        ff = FreeFile
        Open FILEname For Output As #ff
        Close #ff
        
        '//Rename the file with *random* names
        FILEname = RenameFile(FILEname, 5)
        
        If debug_no_delete = False Then
          '//And now a simple kill statement
          DeleteFile FILEname
        End If
        
        '//Update our TOTAL file percentages
        With frmWiping
          .pbTotal.value = Int((FileCount / TotalFiles) * 100)
          .lblTotalPercent = Int((FileCount / TotalFiles) * 100) & " %"
        End With
        
        '//Remove the file from the LV control
        lvFiles.ListItems.Remove (1)
        
        '//Update number of files in Status Bar
        sb.Panels.Item(2).Text = "Files: " & Trim$(Str$(lvFiles.ListItems.Count))
      
        lvFiles.Refresh
        sb.Refresh
        frmMain.Refresh
        DoEvents
      Loop
      
      '//Process the empty root directories now...
      For i = 0 To lstDirectories.ListCount - 1
        FILEname = lstDirectories.List(i)
        File1.Path = FILEname
        If File1.ListCount = 0 Then
          If Wipe_In_Progress = True Then
            '//Rename the directories with *random* names
            FILEname = RenameFile(FILEname, 5)
            If debug_no_delete = False Then
              DeleteDirectory FILEname
            End If
          End If
        End If
      Next i
      
      '//Stop our benchmark timer
      EndTime! = Timer - StartTime!
  
      Wipe_In_Progress = False
        
      '//Set our controls back to normal pre-wipe state
      frmMain.MousePointer = vbDefault
        
      ShowMessage TotalFiles & " Files successfully Wiped in " & Round(EndTime!, 1) & _
        " Seconds." & vbCrLf & "Total Bytes Written: " & TotalBytes, App.Title
      
      '//Reset the directory list array
      If lvFiles.ListItems.Count = 0 Then
        lstDirectories.Clear
      End If
      
      '//Refresh the list box control to start state ONLY if all files
      '//have been deleted.  Otherwise we will just leave the TV as is.
      If lvFiles.ListItems.Count = 0 Then
        AddHelpMSGtoList
      End If

      '//
      Unload frmWiping
      Me.Show

    Case 1
      '//Clear the list view of all files
      ClearAllListItems
      lstDirectories.Clear
      AddHelpMSGtoList
  
    Case 2    '//Settings
      Load frmOptions
      frmOptions.Show , frmMain
  
    Case 3    '//Help
      Dim ret As Long
      ret = ShellExecute(Me.hwnd, vbNullString, "ReadMe.txt", vbNullString, App.Path & "\DOC", SW_SHOWNORMAL)
  End Select
End Sub

Private Sub lvFiles_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 46     '// DEL key
      '//Remove the selected item from the list
      RemoveItemFromList lvFiles.SelectedItem.Index
  End Select
End Sub

Private Sub lvFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Button
    Case 1
    Case 2    '//Right Click
      PopupMenu mnulvPopUp
    Case 3
  End Select
End Sub

Private Sub lvFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  '//Handle DataObjects dragged onto the list box control.
 
  Dim i           As Integer
  Dim numFiles    As Integer
  
  '//Remove our startup message
  ClearAllListItems
  
  numFiles = Data.Files.Count
  For i = 1 To numFiles
    If IsDirectory(Data.Files(i)) = True Then
      '//Add the root directory to list view
      lstDirectories.AddItem Data.Files(i)
      
      GetAllFilesInDir Data.Files(i) & "\*.*"
    Else
      '//It's just a file so we just add it to the list view
      GetFileInformation Data.Files(i)
      AddFileToListView FILE_INFO
    End If
  Next i
  
  '//See if a blank directory tree was dropped (no files)
  If numFiles > 0 And lvFiles.ListItems.Count = 0 Then
    AddHelpMSGtoList
    lstDirectories.Clear
    ShowMessage "Directory contains no files"
  End If
  
  '//Update the file count display in status bar
  sb.Panels.Item(2).Text = "Files: " & Trim$(Str$(lvFiles.ListItems.Count))
End Sub

Private Sub mnuFileExit_Click()
  ExitProgram
End Sub

Private Sub mnuHelpAbout_Click()
  Load frmAbout:  frmAbout.Show
End Sub

Private Sub mnuHelpTopics_Click()
  Dim ret As Long
  ret = ShellExecute(Me.hwnd, vbNullString, "ReadMe.txt", vbNullString, App.Path & "\DOC", SW_SHOWNORMAL)
End Sub

Private Sub mnulvRemoveAll_Click()
  '//Clear the list view of all files
  ClearAllListItems
End Sub

Private Sub mnulvRemoveItem_Click()
  '//Remove the selected item from the list
  RemoveItemFromList lvFiles.SelectedItem.Index
End Sub

Private Sub mnulvWipeAll_Click()
  '//Wipe all files in the list
  imgControl_Click 0
End Sub

Private Sub mnulvWipeSelected_Click()
  '//Wipe only the selected file.  This sub is similar to the imgControl_Click routine
  '//except that here we don't have to worry about processing directories.
  
  Dim FILEname        As String             '//The File Name
  Dim StartTime       As Single             '//Used to time the routine
  Dim EndTime         As Single             '//Used to tim ethe routine
  Dim lBytesWritten   As Long               '//Number of bytes written to file
  Dim TotalFiles      As Long               '//Total number of files to wipe
  Dim TotalBytes      As Long               '//Total number of bytes written
  Dim FileCount       As Long               '//Counter for progress bar
  
  '//See if there is any files to delete.  If not tell the user.
  If lvFiles.ListItems.Item(1) = MSG_STARTUP_HELP Then
    '//User has not selected a file to wipe, so tell them!!
    ShowMessage MSG_NO_FILES_TO_WIPE, "File Not Found"
    Exit Sub
  End If
  
  If MsgBox(MSG_DELETE_THIS_FILE, vbYesNo + vbCritical, "Wipe Files") = vbNo Then
    '//User did not want to commit to the wipe
    Exit Sub
  End If
    
  '//Set flag that we are wiping
  Wipe_In_Progress = True
  
  frmMain.MousePointer = vbHourglass
  DoEvents

  '//Set the start time for benchmarking
  StartTime! = Timer
  
  TotalFiles = 1
  
  Load frmWiping:  frmWiping.Show
      
  FILEname = lvFiles.ListItems.Item(lvFiles.SelectedItem.Index)
  frmWiping.lblFileName.Caption = TruncateFilename(FILEname, 25)

  '// ***** Overwrite the file *****
  If DOD_Pattern <> GUTMANN_PATTERN Then
    '//A DOD pattern is being used
    lBytesWritten = DoD_elete(FILEname, DOD_Pattern)
  Else
    '//The Gutmann patterns are being used
    lBytesWritten = GutmannDelete(FILEname)
  End If
  
  FileCount = FileCount + 1
  TotalBytes = TotalBytes + lBytesWritten
  
  '//Rename the file with *random* names
  FILEname = RenameFile(FILEname, 5)
  
  If debug_no_delete = False Then
    '//And now a simple kill statement
    DeleteFile FILEname
  End If
  
  '//Update the progress bar
  With frmWiping
    .pbTotal.value = Int((FileCount / TotalFiles) * 100)
    .lblTotalPercent = Int((FileCount / TotalFiles) * 100) & " %"
  End With
    
  '//Remove the file from the LV control
  lvFiles.ListItems.Remove (lvFiles.SelectedItem.Index)
  
  '//Update number of files in Status Bar
  sb.Panels.Item(2).Text = "Files: " & Trim$(Str$(lvFiles.ListItems.Count))

  lvFiles.Refresh
  sb.Refresh
  frmMain.Refresh
  DoEvents
      
  '//Stop our benchmark timer
  EndTime! = Timer - StartTime!
  
  Wipe_In_Progress = False
        
  '//Set our controls back to normal pre-wipe state
  frmMain.MousePointer = vbDefault
    
  ShowMessage TotalFiles & " Files successfully Wiped in " & Round(EndTime!, 1) & _
    " Seconds." & vbCrLf & "Total Bytes Written: " & TotalBytes, App.Title
  
  '//Refresh the list box control to start state ONLY if all files
  '//have been deleted.  Otherwise we will just leave the TV as is.
  If lvFiles.ListItems.Count = 0 Then
    AddHelpMSGtoList
  End If
  
  Unload frmWiping
  Me.Show
End Sub

Private Sub mnuOptionsSettings_Click()
  Load frmOptions
  frmOptions.Show
End Sub

Private Sub ClearAllListItems()
  lvFiles.ListItems.Clear
  sb.Panels.Item(2).Text = "Files: " & Trim$(Str$(lvFiles.ListItems.Count))
End Sub

Private Sub RemoveItemFromList(Index As Integer)
  '//Remove the selected item from the list
  If lvFiles.ListItems.Count > 0 Then
    lvFiles.ListItems.Remove Index
    sb.Panels.Item(2).Text = "Files: " & Trim$(Str$(lvFiles.ListItems.Count))
  End If
End Sub

Private Sub AddHelpMSGtoList()
  Dim objLvi As MSComctlLib.ListItem
  
  Set objLvi = frmMain.lvFiles.ListItems.Add()
  objLvi.Text = MSG_STARTUP_HELP
  Set objLvi = Nothing
End Sub

Private Sub ExitProgram()
  Unload Me
  End
End Sub

Private Function GetFileNameFromPath(FILEpath As String) As String
  '//This sub extracts the file name from a full path.  It is used
  '//to build the lstFileNames array control for the registry scanner.
  
  Dim pos As Long
  
  '//Get position of the last "\"
  pos = InStrRev(FILEpath, "\")
  GetFileNameFromPath = Mid$(FILEpath, pos + 1)
End Function

Private Sub LoadRegistrySettings()
  Dim tmp_reg As String
  Dim chk As Long

  '//Make sure registry value is present
  RegCreateRoot
  
  '//---
  '//Get settings from registry (or set defaults)
  '//---
    
  tmp_reg = RegReadOverwritePattern
  Select Case tmp_reg
    Case "0"    '//US DoD 5220.22-M (8-306 /E, C and E)
      DOD_Pattern = 0
    Case "1"    '//US DoD 5220.22-M (8-306 /E)
      DOD_Pattern = 1
    Case "2"    '//Pseudorandom Data
      DOD_Pattern = 2
    Case "3"    '//Gutmann Wipe
      DOD_Pattern = 3
    Case Else
      '//Set Default
      DOD_Pattern = 0
      RegWriteOverwritePattern "0"
  End Select

  tmp_reg = RegReadRenameFiles
  If tmp_reg <> "ERROR" Then
    If tmp_reg = "0" Then
      Rename_Files = False
    Else
      Rename_Files = True
    End If
  Else
    Rename_Files = True
    RegWriteRenameFiles "1"
  End If
End Sub

