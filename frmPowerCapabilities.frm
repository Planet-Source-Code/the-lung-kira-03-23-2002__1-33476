VERSION 5.00
Begin VB.Form frmPowerCapabilities 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Power Capabilities"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6735
   Icon            =   "frmPowerCapabilities.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSoftLidWake 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtRtcWake 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtMinDeviceWakeState 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtDefaultLowLatencyWake 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtAcOnLineWake 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox chkProcessorThrottle 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtBatteryScaleCapacity2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtBatteryScaleGranularity2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtBatteryScaleCapacity1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtBatteryScaleGranularity1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtBatteryScaleCapacity0 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtBatteryScaleGranularity0 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   600
      Width           =   1335
   End
   Begin VB.CheckBox chkShortTermBattery 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   31
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox chkSystemBattery 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   35
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox chkDiskSpinDown 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtProcessorMaxThrottle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtProcessorMinThrottle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.CheckBox chkThermalControl 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   37
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox chkUPS 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   39
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox chkAPM 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkVideoDim 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkFullWake 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkHibernationFile 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkSleepStateS5 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkSleepStateS4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkSleepStateS3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkSleepStateS2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkSleepStateS1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkLidSwitch 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox chkSleepButton 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   33
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox chkPowerButton 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   29
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label lblBatteryScale2 
      Caption         =   "2"
      Height          =   255
      Left            =   3120
      TabIndex        =   49
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblBatteryScale1 
      Caption         =   "1"
      Height          =   255
      Left            =   3120
      TabIndex        =   46
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblBatteryScale0 
      Caption         =   "0"
      Height          =   255
      Left            =   3120
      TabIndex        =   43
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblBatteryScaleCapacity 
      Caption         =   "Capacity"
      Height          =   255
      Left            =   5280
      TabIndex        =   42
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblBatteryScaleGranularity 
      Caption         =   "Granularity"
      Height          =   255
      Left            =   3720
      TabIndex        =   41
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblBatteryScale 
      Caption         =   "Battery Scale"
      Height          =   255
      Left            =   3120
      TabIndex        =   40
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblDefaultLowLatencyWake 
      Caption         =   "Default Low Latency Wake"
      Height          =   255
      Left            =   3120
      TabIndex        =   54
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblMinDeviceWakeState 
      Caption         =   "Min Device Wake State"
      Height          =   255
      Left            =   3120
      TabIndex        =   56
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblRtcWake 
      Caption         =   "Rtc Wake"
      Height          =   255
      Left            =   3120
      TabIndex        =   58
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblSoftLidWake 
      Caption         =   "Soft Lid Wake"
      Height          =   255
      Left            =   3120
      TabIndex        =   60
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblAcOnLineWake 
      Caption         =   "Ac OnLine Wake"
      Height          =   255
      Left            =   3120
      TabIndex        =   52
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblShortTermBattery 
      Caption         =   "Short Term Battery"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblSystemBattery 
      Caption         =   "System Battery"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblDiskSpinDown 
      Caption         =   "Disk Spin Down"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblProcessorMaxThrottle 
      Caption         =   "Processor Maximum Throttle"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblProcessorMinThrottle 
      Caption         =   "Processor Minimum Throttle"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblProcessorThrottle 
      Caption         =   "Processor Throttle"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblThermalControl 
      Caption         =   "Thermal Control"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblUPS 
      Caption         =   "UPS"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblAPM 
      Caption         =   "APM"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblVideoDim 
      Caption         =   "Video Dim"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblFullWake 
      Caption         =   "Full Wake"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblHibernationFile 
      Caption         =   "Hibernation File"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblSleepStateS5 
      Caption         =   "Sleep State S5"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblSleepStateS4 
      Caption         =   "Sleep State S4"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblSleepStateS3 
      Caption         =   "Sleep State S3"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblSleepStateS2 
      Caption         =   "Sleep State S2"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblSleepStateS1 
      Caption         =   "Sleep State S1"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblLidSwitch 
      Caption         =   "Lid Switch"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblSleepButton 
      Caption         =   "Sleep Button"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblPowerButton 
      Caption         =   "Power Button"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3600
      Width           =   2055
   End
End
Attribute VB_Name = "frmPowerCapabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmPowerCapabilities"


Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("powrprof.dll", "CallNtPowerInformation") = True Then
        Dim SYSTEM_POWER_CAPABILITIES As SYSTEM_POWER_CAPABILITIES
        If CallNtPowerInformation(SystemPowerCapabilities, ByVal 0&, 0&, SYSTEM_POWER_CAPABILITIES, Len(SYSTEM_POWER_CAPABILITIES)) <> ERROR_SUCCESS Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "CallNtPowerInformation")
        
        With SYSTEM_POWER_CAPABILITIES
            chkPowerButton.value = IIf(.PowerButtonPresent, 1, 0)
            chkSleepButton.value = IIf(.SleepButtonPresent, 1, 0)
            chkLidSwitch.value = IIf(.LidPresent, 1, 0)
            chkSleepStateS1.value = IIf(.SystemS1, 1, 0)
            chkSleepStateS2.value = IIf(.SystemS2, 1, 0)
            chkSleepStateS3.value = IIf(.SystemS3, 1, 0)
            chkSleepStateS4.value = IIf(.SystemS4, 1, 0)
            chkSleepStateS5.value = IIf(.SystemS5, 1, 0)
            chkHibernationFile.value = IIf(.HiberFilePresent, 1, 0)
            chkFullWake.value = IIf(.FullWake, 1, 0)
            chkVideoDim.value = IIf(.VideoDimPresent, 1, 0)
            chkAPM.value = IIf(.ApmPresent, 1, 0)
            chkUPS.value = IIf(.UpsPresent, 1, 0)
            chkThermalControl.value = IIf(.ThermalControl, 1, 0)
            chkProcessorThrottle.value = IIf(.ProcessorThrottle, 1, 0)
            txtProcessorMinThrottle.Text = .ProcessorMinThrottle & "%"
            txtProcessorMaxThrottle.Text = .ProcessorMaxThrottle & "%"
            chkDiskSpinDown.value = IIf(.DiskSpinDown, 1, 0)
            chkSystemBattery.value = IIf(.SystemBatteriesPresent, 1, 0)
            chkShortTermBattery.value = IIf(.BatteriesAreShortTerm, 1, 0)
            
            txtBatteryScaleGranularity0.Text = .BatteryScale(0).Granularity
            txtBatteryScaleGranularity1.Text = .BatteryScale(1).Granularity
            txtBatteryScaleGranularity2.Text = .BatteryScale(2).Granularity
            txtBatteryScaleCapacity0.Text = .BatteryScale(0).Capacity
            txtBatteryScaleCapacity1.Text = .BatteryScale(1).Capacity
            txtBatteryScaleCapacity2.Text = .BatteryScale(2).Capacity
            
            txtAcOnLineWake.Text = SystemPowerState(.AcOnLineWake)
            txtSoftLidWake.Text = SystemPowerState(.SoftLidWake)
            txtRtcWake.Text = SystemPowerState(.RtcWake)
            txtMinDeviceWakeState.Text = SystemPowerState(.MinDeviceWakeState)
            txtDefaultLowLatencyWake.Text = SystemPowerState(.DefaultLowLatencyWake)
        End With
    Else
        lblPowerButton.Enabled = False
        lblSleepButton.Enabled = False
        lblLidSwitch.Enabled = False
        lblSleepStateS1.Enabled = False
        lblSleepStateS2.Enabled = False
        lblSleepStateS3.Enabled = False
        lblSleepStateS4.Enabled = False
        lblSleepStateS5.Enabled = False
        lblHibernationFile.Enabled = False
        lblFullWake.Enabled = False
        lblVideoDim.Enabled = False
        lblAPM.Enabled = False
        lblUPS.Enabled = False
        lblThermalControl.Enabled = False
        lblProcessorThrottle.Enabled = False
        lblProcessorMinThrottle.Enabled = False
        lblProcessorMaxThrottle.Enabled = False
        lblDiskSpinDown.Enabled = False
        lblSystemBattery.Enabled = False
        lblShortTermBattery.Enabled = False
        lblBatteryScale.Enabled = False
        lblBatteryScaleGranularity.Enabled = False
        lblBatteryScaleCapacity.Enabled = False
        lblBatteryScale0.Enabled = False
        lblBatteryScale1.Enabled = False
        lblBatteryScale2.Enabled = False
        lblAcOnLineWake.Enabled = False
        lblSoftLidWake.Enabled = False
        lblRtcWake.Enabled = False
        lblMinDeviceWakeState.Enabled = False
        lblDefaultLowLatencyWake.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
