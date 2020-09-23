VERSION 5.00
Begin VB.Form frmBatteryState 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Battery State"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "frmBatteryState.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWarningAlert 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtLowAlert 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "S"
      Top             =   2160
      Width           =   135
   End
   Begin VB.TextBox txtTimeRemaining 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtmWh3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "mWh"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtDischargeRate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtmWh2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "mWh"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtRemainingCapacity 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtmWh1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "mWh"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtMaxCapacity 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CheckBox chkDischarging 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox chkCharging 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkBatteryPresent 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkACOnline 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblWarningAlert 
      Caption         =   "Warning Alert"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblLowAlert 
      Caption         =   "Low Alert"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblTimeRemaining 
      Caption         =   "Time Remaining"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblDischargeRate 
      Caption         =   "Discharge Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblRemainingCapacity 
      Caption         =   "Remaining Capacity"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblMaxCapacity 
      Caption         =   "Max Capacity"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblDischarging 
      Caption         =   "Discharging"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblCharging 
      Caption         =   "Charging"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblBatteryPresent 
      Caption         =   "Battery Present"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblACOnline 
      Caption         =   "AC Online"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmBatteryState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmBatteryState"


Private Sub Form_Load()
On Error GoTo VB_Error
    
    If Function_Exist("powrprof.dll", "CallNtPowerInformation") = True Then
        Dim SYSTEM_BATTERY_STATE As SYSTEM_BATTERY_STATE
        If CallNtPowerInformation(SystemBatteryState, ByVal 0&, 0&, SYSTEM_BATTERY_STATE, Len(SYSTEM_BATTERY_STATE)) <> ERROR_SUCCESS Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "CallNtPowerInformation")
        
        With SYSTEM_BATTERY_STATE
            chkACOnline.value = IIf(.AcOnLine, 1, 0)
            chkBatteryPresent.value = IIf(.BatteryPresent, 1, 0)
            chkCharging.value = IIf(.Charging, 1, 0)
            txtLowAlert.Text = FormatNumber(int32_uint32(.DefaultAlert1), 0, , , True)
            txtWarningAlert.Text = FormatNumber(int32_uint32(.DefaultAlert2), 0, , , True)
            chkDischarging.value = IIf(.Discharging, 1, 0)
            txtTimeRemaining.Text = FormatNumber(int32_uint32(.EstimatedTime), 0, , , True)
            txtMaxCapacity.Text = FormatNumber(int32_uint32(.MaxCapacity), 0, , , True)
            txtRemainingCapacity.Text = FormatNumber(int32_uint32(.RemainingCapacity), 0, , , True)
            txtDischargeRate.Text = FormatNumber(int32_uint32(.Rate), 0, , , True)
        End With
    Else
        lblACOnline.Enabled = False
        lblBatteryPresent.Enabled = False
        lblCharging.Enabled = False
        lblLowAlert.Enabled = False
        lblWarningAlert.Enabled = False
        lblDischarging.Enabled = False
        lblTimeRemaining.Enabled = False
        lblMaxCapacity.Enabled = False
        lblRemainingCapacity.Enabled = False
        lblDischargeRate.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
