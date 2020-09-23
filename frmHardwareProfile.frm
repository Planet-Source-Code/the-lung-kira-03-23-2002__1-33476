VERSION 5.00
Begin VB.Form frmHardwareProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hardware Profile"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmHardwareProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDockingInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtGUID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblDockingInfo 
      Caption         =   "Docking Info"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblGUID 
      Caption         =   "GUID"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmHardwareProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmHardwareProfile"


Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("advapi32.dll", "GetCurrentHwProfileA") = True Then
        Dim HW_PROFILE_INFO As HW_PROFILE_INFO
        If GetCurrentHwProfile(HW_PROFILE_INFO) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "GetCurrentHwProfile")
        
        With HW_PROFILE_INFO
            txtName.Text = .szHwProfileName
            txtGUID.Text = .szHwProfileGuid
            
            If DOCKINFO_DOCKED And .dwDockInfo Then txtDockingInfo.Text = "Docked"
            If DOCKINFO_UNDOCKED And .dwDockInfo Then txtDockingInfo.Text = "Undocked"
            If DOCKINFO_USER_SUPPLIED And .dwDockInfo Then txtDockingInfo.Text = txtDockingInfo.Text & " UserSupplied"
        End With
    Else
        lblName.Enabled = False
        lblGUID.Enabled = False
        lblDockingInfo.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
