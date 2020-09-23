VERSION 5.00
Begin VB.Form frmProcessorPowerInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processor Power Info"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmProcessorPowerInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMhzLimit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtMaxMhz 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtMaxIdleState 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtCurrentIdleState 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtCurrentMhz 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cboProcessor 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCurrentIdleState 
      Caption         =   "Current Idle State"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblMaxIdleState 
      Caption         =   "Max Idle State"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblMhzLimit 
      Caption         =   "Mhz Limit"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblCurrentMhz 
      Caption         =   "Current Mhz"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblMaxMhz 
      Caption         =   "Max Mhz"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblProcessor 
      Caption         =   "Processor"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmProcessorPowerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PROCESSOR_POWER_INFORMATION() As PROCESSOR_POWER_INFORMATION
Const sLocation As String = "frmProcessorPowerInfo"


Private Sub cboProcessor_Click()
On Error GoTo VB_Error

    With PROCESSOR_POWER_INFORMATION(cboProcessor.ListIndex)
        txtMaxMhz.Text = FormatNumber(int32_uint32(.MaxMhz), 0, , , True)
        txtCurrentMhz.Text = FormatNumber(int32_uint32(.CurrentMhz), 0, , , True)
        txtMhzLimit.Text = FormatNumber(int32_uint32(.MhzLimit), 0, , , True)
        txtMaxIdleState.Text = int32_uint32(.MaxIdleState)
        txtCurrentIdleState.Text = int32_uint32(.CurrentIdleState)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboProcessor_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("powrprof.dll", "CallNtPowerInformation") = True Then
        Dim SYSTEM_INFO As SYSTEM_INFO
        Call GetSystemInfo(SYSTEM_INFO)
        
        If SYSTEM_INFO.dwNumberOrfProcessors > 0 Then
            ReDim PROCESSOR_POWER_INFORMATION(SYSTEM_INFO.dwNumberOrfProcessors - 1) As PROCESSOR_POWER_INFORMATION
            If CallNtPowerInformation(ProcessorInformation, ByVal 0&, 0&, PROCESSOR_POWER_INFORMATION(0), Len(PROCESSOR_POWER_INFORMATION(0)) * SYSTEM_INFO.dwNumberOrfProcessors) <> ERROR_SUCCESS Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "CallNtPowerInformation")
            
            Dim lIncrement As Long
            For lIncrement = 0 To (SYSTEM_INFO.dwNumberOrfProcessors - 1)
                cboProcessor.AddItem lIncrement
            Next lIncrement
        End If
    Else
        lblProcessor.Enabled = False
        cboProcessor.Enabled = False
        lblMaxMhz.Enabled = False
        lblCurrentMhz.Enabled = False
        lblMhzLimit.Enabled = False
        lblMaxIdleState.Enabled = False
        lblCurrentIdleState.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
