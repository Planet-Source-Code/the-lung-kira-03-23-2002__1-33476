VERSION 5.00
Begin VB.Form frmCPUInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPU Info"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmCPUInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLowEndCPU 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtMS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3960
      TabIndex        =   8
      Text            =   "MHz"
      Top             =   960
      Width           =   375
   End
   Begin VB.Timer tmrCyclesElapsed 
      Interval        =   1000
      Left            =   1680
      Top             =   1080
   End
   Begin VB.TextBox txtSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtArchitecture 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtActiveProcessorMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtProcessors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtCyclesElapsed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblActiveProcessorMask 
      Caption         =   "Active Processor Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLowEndCPU 
      Caption         =   "Low End CPU"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblProcessors 
      Caption         =   "Number of Processors"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Approx Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblCyclesElapsed 
      Caption         =   "Cycles Elapsed"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblArchitecture 
      Caption         =   "Architecture"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmCPUInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmCPUInfo"


Private Sub Form_Activate()
On Error GoTo VB_Error

    DoEvents
    
    
    Dim dCyclesB As Double
    Dim dCyclesA
    Dim dTime As Double
    
    dTime = PerformanceCounter
    dCyclesB = rdtsc_
    
    Do
    Loop While PerformanceCounter < (dTime + dCounterFrequency)
    
    dCyclesA = rdtsc_
    
    txtSpeed.Text = FormatNumber((dCyclesA - dCyclesB) / 1000000, 0, , , True)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Activate")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim SYSTEM_INFO As SYSTEM_INFO
    GetSystemInfo SYSTEM_INFO
    
    txtActiveProcessorMask.Text = StrReverse(Right$(String$(32, "0") & ltoa_(SYSTEM_INFO.dwActiveProcessorMask, 2), 32))
    txtCyclesElapsed.Text = FormatNumber$(rdtsc_, 0, , , True)
    
    Select Case LOWORD(SYSTEM_INFO.dwOemID)
        Case PROCESSOR_ARCHITECTURE_INTEL: txtArchitecture.Text = "Intel"
        Case PROCESSOR_ARCHITECTURE_MIPS: txtArchitecture.Text = "MIPS"
        Case PROCESSOR_ARCHITECTURE_ALPHA: txtArchitecture.Text = "Alpha"
        Case PROCESSOR_ARCHITECTURE_PPC: txtArchitecture.Text = "PPC"
        Case PROCESSOR_ARCHITECTURE_SHX: txtArchitecture.Text = "SHX"
        Case PROCESSOR_ARCHITECTURE_ARM: txtArchitecture.Text = "ARM"
        Case PROCESSOR_ARCHITECTURE_IA64: txtArchitecture.Text = "IA-64"
        Case PROCESSOR_ARCHITECTURE_ALPHA64: txtArchitecture.Text = "Alpha 64"
        Case PROCESSOR_ARCHITECTURE_MSIL: txtArchitecture.Text = "MSIL"
        Case PROCESSOR_ARCHITECTURE_AMD64: txtArchitecture.Text = "AMD 64"
        Case PROCESSOR_ARCHITECTURE_IA32_ON_WIN64: txtArchitecture.Text = "IA32 On Win64"
        Case PROCESSOR_ARCHITECTURE_UNKNOWN: txtArchitecture.Text = "Unknown"
        Case Else: txtArchitecture.Text = "Unknown " & LOWORD(SYSTEM_INFO.dwOemID)
    End Select
    
    txtProcessors.Text = int32_uint32(SYSTEM_INFO.dwNumberOrfProcessors)
    chkLowEndCPU.value = IIf(GetSystemMetrics(SM_SLOWMACHINE), 1, 0)
    
    
    tmrCyclesElapsed_Timer
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    tmrCyclesElapsed.Enabled = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub tmrCyclesElapsed_Timer()
On Error GoTo VB_Error

    txtCyclesElapsed.Text = FormatNumber$(rdtsc_, 0, , , True)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\tmrCyclesElapsed_Timer")
Resume Next
End Sub
