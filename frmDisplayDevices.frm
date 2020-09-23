VERSION 5.00
Begin VB.Form frmDisplayDevices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Devices"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmDisplayDevices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkVGACompatible 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chkRemovable 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkPrimaryDevice 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   2400
      Width           =   255
   End
   Begin VB.CheckBox chkModesPruned 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chkMirroringDriver 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox chkAttachedToDesktop 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtDeviceName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtDeviceKey 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtDeviceID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.ComboBox cboDisplayDevices 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label lblVGACompatible 
      Caption         =   "VGA Compatible"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblRemovable 
      Caption         =   "Removable"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblPrimaryDevice 
      Caption         =   "Primary Device"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblModesPruned 
      Caption         =   "Modes Pruned"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblMirroringDriver 
      Caption         =   "Mirroring Driver"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblAttachedToDesktop 
      Caption         =   "Attached To Desktop"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblDeviceKey 
      Caption         =   "Device Key"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblDeviceID 
      Caption         =   "Device ID"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblDeviceName 
      Caption         =   "Device Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblDisplayDevices 
      Caption         =   "Display Devices"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmDisplayDevices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim DISPLAY_DEVICE() As DISPLAY_DEVICE
Const sLocation As String = "frmDisplayDevices"


Private Sub cboDisplayDevices_Click()
On Error GoTo VB_Error

    With DISPLAY_DEVICE(cboDisplayDevices.ListIndex)
        txtDeviceID.Text = .DeviceID
        txtDeviceKey.Text = .DeviceKey
        txtDeviceName.Text = .DeviceName
        
        chkAttachedToDesktop.value = IIf(.StateFlags And DISPLAY_DEVICE_ATTACHED_TO_DESKTOP, 1, 0)
        chkMirroringDriver.value = IIf(.StateFlags And DISPLAY_DEVICE_MIRRORING_DRIVER, 1, 0)
        chkModesPruned.value = IIf(.StateFlags And DISPLAY_DEVICE_MODESPRUNED, 1, 0)
        chkPrimaryDevice.value = IIf(.StateFlags And DISPLAY_DEVICE_PRIMARY_DEVICE, 1, 0)
        chkRemovable.value = IIf(.StateFlags And DISPLAY_DEVICE_REMOVABLE, 1, 0)
        chkVGACompatible.value = IIf(.StateFlags And DISPLAY_DEVICE_VGA_COMPATIBLE, 1, 0)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboDisplayDevices_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("user32.dll", "EnumDisplayDevicesA") = True Then
        Dim lDeviceNumber As Long
        
        Do
            ReDim Preserve DISPLAY_DEVICE(lDeviceNumber)
            DISPLAY_DEVICE(lDeviceNumber).cb = Len(DISPLAY_DEVICE(lDeviceNumber))
            If EnumDisplayDevices(0&, lDeviceNumber, DISPLAY_DEVICE(lDeviceNumber), 0&) = False Then
                If Err.LastDllError <> 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "EnumDisplayDevices")
                Exit Do
            Else
                cboDisplayDevices.AddItem DISPLAY_DEVICE(lDeviceNumber).DeviceString
                lDeviceNumber = lDeviceNumber + 1
            End If
            
            If bShutdown = True Then Exit Do
        Loop
        
        If cboDisplayDevices.ListCount > 0 Then cboDisplayDevices.ListIndex = 0
    Else
        lblDisplayDevices.Enabled = False
        cboDisplayDevices.Enabled = False
        lblDeviceID.Enabled = False
        lblDeviceKey.Enabled = False
        lblDeviceName.Enabled = False
        lblAttachedToDesktop.Enabled = False
        lblMirroringDriver.Enabled = False
        lblModesPruned.Enabled = False
        lblPrimaryDevice.Enabled = False
        lblRemovable.Enabled = False
        lblVGACompatible.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
