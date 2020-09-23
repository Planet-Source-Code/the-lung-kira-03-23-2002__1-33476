VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDirectories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directories"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmDirectories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwDirectories 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmDirectories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmDirectories"


Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwDirectories.ColumnHeaders
        .Add , , "Description"
        .Add , , "Path"
    End With
    
    If Function_Exist("shell32.dll", "SHGetSpecialFolderPathA") = True Then
        With lvwDirectories.ListItems
            .Add(, , "Admin Tools").SubItems(1) = Dir_Path_Get(CSIDL_ADMINTOOLS)
            .Add(, , "Alt Startup").SubItems(1) = Dir_Path_Get(CSIDL_ALTSTARTUP)
            .Add(, , "App Data").SubItems(1) = Dir_Path_Get(CSIDL_APPDATA)
            .Add(, , "Common Admin Tools").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_ADMINTOOLS)
            .Add(, , "Common Alt Startup").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_ALTSTARTUP)
            .Add(, , "Common App Data").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_APPDATA)
            .Add(, , "Common Desktop").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_DESKTOPDIRECTORY)
            .Add(, , "Common Documents").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_DOCUMENTS)
            .Add(, , "Common Favorites").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_FAVORITES)
            .Add(, , "Common Program Files").SubItems(1) = Dir_Path_Get(CSIDL_PROGRAM_FILES_COMMON)
            .Add(, , "Common Programs").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_PROGRAMS)
            .Add(, , "Common StartMenu").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_STARTMENU)
            .Add(, , "Common Startup").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_STARTUP)
            .Add(, , "Common Templates").SubItems(1) = Dir_Path_Get(CSIDL_COMMON_TEMPLATES)
            .Add(, , "Controls").SubItems(1) = Dir_Path_Get(CSIDL_CONTROLS)
            .Add(, , "Cookies").SubItems(1) = Dir_Path_Get(CSIDL_COOKIES)
            .Add(, , "Desktop").SubItems(1) = Dir_Path_Get(CSIDL_DESKTOP)
            .Add(, , "Desktop Directory").SubItems(1) = Dir_Path_Get(CSIDL_DESKTOPDIRECTORY)
            .Add(, , "Drives").SubItems(1) = Dir_Path_Get(CSIDL_DRIVES)
            .Add(, , "Favorites").SubItems(1) = Dir_Path_Get(CSIDL_FAVORITES)
            .Add(, , "Fonts").SubItems(1) = Dir_Path_Get(CSIDL_FONTS)
            .Add(, , "History").SubItems(1) = Dir_Path_Get(CSIDL_HISTORY)
            .Add(, , "Internet").SubItems(1) = Dir_Path_Get(CSIDL_INTERNET)
            .Add(, , "Internet Cache").SubItems(1) = Dir_Path_Get(CSIDL_INTERNET_CACHE)
            .Add(, , "Local App Data").SubItems(1) = Dir_Path_Get(CSIDL_LOCAL_APPDATA)
            .Add(, , "My Documents").SubItems(1) = Dir_Path_Get(CSIDL_PERSONAL)
            .Add(, , "My Network Places").SubItems(1) = Dir_Path_Get(CSIDL_NETHOOD)
            .Add(, , "My Pictures").SubItems(1) = Dir_Path_Get(CSIDL_MYPICTURES)
            .Add(, , "Network Neighborhood").SubItems(1) = Dir_Path_Get(CSIDL_NETWORK)
            .Add(, , "Printers").SubItems(1) = Dir_Path_Get(CSIDL_PRINTERS)
            .Add(, , "Print Hood").SubItems(1) = Dir_Path_Get(CSIDL_PRINTHOOD)
            .Add(, , "Profile").SubItems(1) = Dir_Path_Get(CSIDL_PROFILE)
            .Add(, , "Program Files").SubItems(1) = Dir_Path_Get(CSIDL_PROGRAM_FILES)
            .Add(, , "Programs").SubItems(1) = Dir_Path_Get(CSIDL_PROGRAMS)
            .Add(, , "Recent").SubItems(1) = Dir_Path_Get(CSIDL_RECENT)
            .Add(, , "RecyleBin").SubItems(1) = Dir_Path_Get(CSIDL_BITBUCKET)
            .Add(, , "SendTo").SubItems(1) = Dir_Path_Get(CSIDL_SENDTO)
            .Add(, , "StartMenu").SubItems(1) = Dir_Path_Get(CSIDL_STARTMENU)
            .Add(, , "Startup").SubItems(1) = Dir_Path_Get(CSIDL_STARTUP)
            .Add(, , "System").SubItems(1) = Dir_Path_Get(CSIDL_SYSTEM)
            .Add(, , "Templates").SubItems(1) = Dir_Path_Get(CSIDL_TEMPLATES)
            .Add(, , "Windows").SubItems(1) = Dir_Path_Get(CSIDL_WINDOWS)
        End With
    Else
        lvwDirectories.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
