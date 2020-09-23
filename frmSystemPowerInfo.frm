VERSION 5.00
Begin VB.Form frmSystemPowerInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Power Info"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmSystemPowerInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "S"
      Top             =   1080
      Width           =   135
   End
   Begin VB.TextBox txtMaxIdlenessAllowed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtCoolingMode 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtTimeRemaining 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtCurrentIdleness 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCoolingMode 
      Caption         =   "Cooling Mode"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblTimeRemaining 
      Caption         =   "Time Remaining"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblCurrentIdleness 
      Caption         =   "Current Idleness"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblEstablishedConnections 
      Caption         =   "Max Idleness Allowed"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmSystemPowerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmSystemPowerInfo"


Private Sub Form_Load()
On Error GoTo VB_Error
    
    If Function_Exist("powrprof.dll", "CallNtPowerInformation") = True Then
        Dim SYSTEM_POWER_INFORMATION As SYSTEM_POWER_INFORMATION
        If CallNtPowerInformation(SystemPowerInformation, ByVal 0&, 0&, SYSTEM_POWER_INFORMATION, Len(SYSTEM_POWER_INFORMATION)) <> ERROR_SUCCESS Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "CallNtPowerInformation")
        
        With SYSTEM_POWER_INFORMATION
            txtMaxIdlenessAllowed.Text = int32_uint32(.MaxIdlenessAllowed) & "%"
            txtCurrentIdleness.Text = int32_uint32(.Idleness) & "%"
            txtTimeRemaining.Text = FormatNumber(int32_uint32(.TimeRemaining), 0, , , True)
            
            Select Case .CoolingMode
                'Case PO_TZ_ACTIVE: txtCoolingMode.Text = "Active"
                'Case PO_TZ_INVALID_MODE: txtCoolingMode.Text = "None"
                'Case PO_TZ_PASSIVE: txtCoolingMode.Text = "Passive"
                Case Else: txtCoolingMode.Text = .CoolingMode
            End Select
        End With
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
