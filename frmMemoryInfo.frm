VERSION 5.00
Begin VB.Form frmMemoryInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Info"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmMemoryInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtB 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "B"
      Top             =   840
      Width           =   135
   End
   Begin VB.TextBox txtPageSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtMinimumApplicationAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtMaximumApplicationAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtAllocationGranularity 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblAllocationGranularity 
      Caption         =   "Allocation Granularity"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblMaximumApplicationAddress 
      Caption         =   "Maximum Application Address"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblMinimumApplicationAddress 
      Caption         =   "Minimum Application Address"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblPageSize 
      Caption         =   "Page Size"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "frmMemoryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMemoryInfo"


Private Sub Form_Load()
On Error GoTo VB_Error
    
    Dim SYSTEM_INFO As SYSTEM_INFO
    Call GetSystemInfo(SYSTEM_INFO)
    
    With SYSTEM_INFO
        txtAllocationGranularity.Text = FormatNumber$(int32_uint32(.dwAllocationGranularity), 0, , , True)
        txtMaximumApplicationAddress.Text = FormatNumber$(int32_uint32(.lpMaximumApplicationAddress), 0, , , True)
        txtMinimumApplicationAddress.Text = FormatNumber$(int32_uint32(.lpMinimumApplicationAddress), 0, , , True)
        txtPageSize.Text = FormatNumber$(int32_uint32(.dwPageSize), 0, , , True)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
