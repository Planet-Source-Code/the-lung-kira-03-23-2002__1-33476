VERSION 5.00
Begin VB.Form frmPowerStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Power Status"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmPowerStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBatteryLifePercent 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Timer tmrPowerStatus 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   480
   End
   Begin VB.TextBox txtBatteryChargeStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtBatteryLifeTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtBatteryFullLifeTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtACLineStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblBatteryFullLifeTime 
      Caption         =   "Battery Full Life"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblBatteryLifeTime 
      Caption         =   "Battery Life"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblBatteryChargeStatus 
      Caption         =   "Battery Charge Status"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblACLineStatus 
      Caption         =   "AC Power Status"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmPowerStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmPowerStatus"


Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("kernel32.dll", "GetSystemPowerStatus") = True Then
        tmrPowerStatus.Enabled = True
        tmrPowerStatus_Timer
    Else
        lblACLineStatus.Enabled = False
        lblBatteryChargeStatus.Enabled = False
        lblBatteryFullLifeTime.Enabled = False
        lblBatteryLifeTime.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    tmrPowerStatus.Enabled = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub tmrPowerStatus_Timer()
On Error GoTo VB_Error

    Dim SYSTEM_POWER_STATUS As SYSTEM_POWER_STATUS
    
    If GetSystemPowerStatus(SYSTEM_POWER_STATUS) = False Then
        tmrPowerStatus.Enabled = False
        Call Error_API(Err.LastDllError, sLocation & "\tmrPowerStatus_Timer", "GetSystemPowerStatus")
    End If
    
    
    With SYSTEM_POWER_STATUS
        Select Case .ACLineStatus
            Case 0: txtACLineStatus.Text = "Offline"
            Case 1: txtACLineStatus.Text = "Online"
            Case 255: txtACLineStatus.Text = "Unkown status"
            Case Else: txtACLineStatus.Text = "Unknown " & .ACLineStatus
        End Select
        
        Select Case .BatteryFlag
            Case 1: txtBatteryChargeStatus.Text = "High"
            Case 2: txtBatteryChargeStatus.Text = "Low"
            Case 4: txtBatteryChargeStatus.Text = "Critical"
            Case 8: txtBatteryChargeStatus.Text = "Charging"
            Case 128: txtBatteryChargeStatus.Text = "No system battery"
            Case 255: txtBatteryChargeStatus.Text = "Unknown status"
            Case Else: txtBatteryChargeStatus.Text = "Unknown " & .BatteryFlag
        End Select
        
        If .BatteryLifePercent <> 255 Then
            txtBatteryLifePercent.Text = .BatteryLifePercent & "%"
        Else
            txtBatteryLifePercent.Text = "?"
        End If
        
        If .BatteryLifeTime > -1 Then
            txtBatteryLifeTime.Text = .BatteryLifeTime
        Else
            txtBatteryLifeTime.Text = "Unknown"
        End If
        If .BatteryFullLifeTime > -1 Then
            txtBatteryFullLifeTime.Text = .BatteryFullLifeTime
        Else
            txtBatteryFullLifeTime.Text = "Unknown"
        End If
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\tmrPowerStatus_Timer")
Resume Next
End Sub
