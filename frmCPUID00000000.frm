VERSION 5.00
Begin VB.Form frmCPUID00000000 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPUID 00000000"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmCPUID00000000.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVendorIDString 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtEDX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtECX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtEBX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtEAX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtMaxCPUIDLevel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblECX 
      Caption         =   "ECX"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblEDX 
      Caption         =   "EDX"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblEAX 
      Caption         =   "EAX"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblEBX 
      Caption         =   "EBX"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblMaxCPUIDLevel 
      Caption         =   "Max CPUID Level"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblVendorIDString 
      Caption         =   "Vendor ID String"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmCPUID00000000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmCPUID00000000"


Private Sub Form_Load()
On Error GoTo VB_Error

    'EAX = 0
    
    Dim lIncrement As Long
    Dim sRegister As String
    Dim tsRegister As String
    
    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    
    cpuid_ 0, outEAX, outEBX, outECX, outEDX
    
    txtMaxCPUIDLevel.Text = Right$("00000000" & ltoa_(outEAX, 16), 8)
    
    sRegister = Right$("00000000" & ltoa_(outEBX, 16), 8)
    For lIncrement = 1 To Len(sRegister)
        tsRegister = tsRegister & Chr$(strtol_(Mid$(sRegister, lIncrement, 2), 16))
        lIncrement = lIncrement + 1
    Next lIncrement
    txtVendorIDString.Text = StrReverse(tsRegister)
    
    sRegister = Right$("00000000" & ltoa_(outEDX, 16), 8)
    tsRegister = vbNullString
    For lIncrement = 1 To Len(sRegister)
        tsRegister = tsRegister & Chr$(strtol_(Mid$(sRegister, lIncrement, 2), 16))
        lIncrement = lIncrement + 1
    Next lIncrement
    txtVendorIDString.Text = txtVendorIDString.Text & StrReverse(tsRegister)
    
    sRegister = Right$("00000000" & ltoa_(outECX, 16), 8)
    tsRegister = vbNullString
    For lIncrement = 1 To Len(sRegister)
        tsRegister = tsRegister & Chr$(strtol_(Mid$(sRegister, lIncrement, 2), 16))
        lIncrement = lIncrement + 1
    Next lIncrement
    txtVendorIDString.Text = txtVendorIDString.Text & StrReverse(tsRegister)
    
    
    txtEAX.Text = Right$("00000000" & ltoa_(outEAX, 16), 8)
    txtEBX.Text = Right$("00000000" & ltoa_(outEBX, 16), 8)
    txtECX.Text = Right$("00000000" & ltoa_(outECX, 16), 8)
    txtEDX.Text = Right$("00000000" & ltoa_(outEDX, 16), 8)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
