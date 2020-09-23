VERSION 5.00
Begin VB.Form frmAdminPowerPolicy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Power Policy"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmAdminPowerPolicy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboMaxSpindownTimeout 
      Height          =   315
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      Width           =   1815
   End
   Begin VB.ComboBox cboMinSpindownTimeout 
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox cboMaxVideoTimeout 
      Height          =   315
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox cboMinVideoTimeout 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox cboMaxSystemPowerSleepState 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox cboMinSystemPowerSleepState 
      Height          =   315
      ItemData        =   "frmAdminPowerPolicy.frx":000C
      Left            =   2640
      List            =   "frmAdminPowerPolicy.frx":000E
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3480
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblMaxSpindownTimeout 
      Caption         =   "Max Spindown Timeout"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label lblMinSpindownTimeout 
      Caption         =   "Min Spindown Timeout"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lblMaxVideoTimeout 
      Caption         =   "Max Video Timeout"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblMinVideoTimeout 
      Caption         =   "Min Video Timeout"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblMaxSystemPowerSleepState 
      Caption         =   "Max System Power  Sleep State"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblMinSystemPowerSleepState 
      Caption         =   "Min System Power  Sleep State"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmAdminPowerPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmAdminPowerPolicy"


Private Sub cmdApply_Click()
On Error GoTo VB_Error

    Dim ADMINISTRATOR_POWER_POLICY As ADMINISTRATOR_POWER_POLICY
    With ADMINISTRATOR_POWER_POLICY
        .MinSleep = cboMinSystemPowerSleepState.ListIndex
        .MaxSleep = cboMaxSystemPowerSleepState.ListIndex
        .MinVideoTimeout = cboMinVideoTimeout.ListIndex
        .MaxVideoTimeout = cboMaxVideoTimeout.ListIndex
        .MinSpindownTimeout = cboMinSpindownTimeout.ListIndex
        .MaxSpindownTimeout = cboMaxSpindownTimeout.ListIndex
    End With
    
    If CallNtPowerInformation(AdministratorPowerPolicy, ADMINISTRATOR_POWER_POLICY, Len(ADMINISTRATOR_POWER_POLICY), ByVal 0&, 0&) <> ERROR_SUCCESS Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "CallNtPowerInformation")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With cboMinSystemPowerSleepState
        .AddItem "Unspecified"
        .AddItem "Working"
        .AddItem "Sleeping1"
        .AddItem "Sleeping2"
        .AddItem "Sleeping3"
        .AddItem "Hibernate"
        .AddItem "Shutdown"
    End With
    With cboMaxSystemPowerSleepState
        .AddItem "Unspecified"
        .AddItem "Working"
        .AddItem "Sleeping1"
        .AddItem "Sleeping2"
        .AddItem "Sleeping3"
        .AddItem "Hibernate"
        .AddItem "Shutdown"
    End With
    With cboMinVideoTimeout
        .AddItem "Unspecified"
        .AddItem "Working"
        .AddItem "Sleeping1"
        .AddItem "Sleeping2"
        .AddItem "Sleeping3"
        .AddItem "Hibernate"
        .AddItem "Shutdown"
    End With
    With cboMaxVideoTimeout
        .AddItem "Unspecified"
        .AddItem "Working"
        .AddItem "Sleeping1"
        .AddItem "Sleeping2"
        .AddItem "Sleeping3"
        .AddItem "Hibernate"
        .AddItem "Shutdown"
    End With
    With cboMinSpindownTimeout
        .AddItem "Unspecified"
        .AddItem "Working"
        .AddItem "Sleeping1"
        .AddItem "Sleeping2"
        .AddItem "Sleeping3"
        .AddItem "Hibernate"
        .AddItem "Shutdown"
    End With
    With cboMaxSpindownTimeout
        .AddItem "Unspecified"
        .AddItem "Working"
        .AddItem "Sleeping1"
        .AddItem "Sleeping2"
        .AddItem "Sleeping3"
        .AddItem "Hibernate"
        .AddItem "Shutdown"
    End With
    
    
    If Function_Exist("powrprof.dll", "CallNtPowerInformation") = True Then
        Dim ADMINISTRATOR_POWER_POLICY As ADMINISTRATOR_POWER_POLICY
        If CallNtPowerInformation(AdministratorPowerPolicy, ByVal 0&, 0&, ADMINISTRATOR_POWER_POLICY, Len(ADMINISTRATOR_POWER_POLICY)) <> ERROR_SUCCESS Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "CallNtPowerInformation")
        
        With ADMINISTRATOR_POWER_POLICY
            cboMinSystemPowerSleepState.ListIndex = MinMax(.MinSleep, 0, 6)
            cboMaxSystemPowerSleepState.ListIndex = MinMax(.MaxSleep, 0, 6)
            cboMinVideoTimeout.ListIndex = MinMax(.MinVideoTimeout, 0, 6)
            cboMaxVideoTimeout.ListIndex = MinMax(.MaxVideoTimeout, 0, 6)
            cboMinSpindownTimeout.ListIndex = MinMax(.MinSpindownTimeout, 0, 6)
            cboMaxSpindownTimeout.ListIndex = MinMax(.MaxSpindownTimeout, 0, 6)
        End With
    Else
        lblMinSystemPowerSleepState.Enabled = False
        cboMinSystemPowerSleepState.Enabled = False
        lblMaxSystemPowerSleepState.Enabled = False
        cboMaxSystemPowerSleepState.Enabled = False
        lblMinVideoTimeout.Enabled = False
        cboMinVideoTimeout.Enabled = False
        lblMaxVideoTimeout.Enabled = False
        cboMaxVideoTimeout.Enabled = False
        lblMinSpindownTimeout.Enabled = False
        cboMinSpindownTimeout.Enabled = False
        lblMaxSpindownTimeout.Enabled = False
        cboMaxSpindownTimeout.Enabled = False
        cmdApply.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
