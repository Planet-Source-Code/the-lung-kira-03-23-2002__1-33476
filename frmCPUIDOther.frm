VERSION 5.00
Begin VB.Form frmCPUIDOther 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPUID Other"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmCPUIDOther.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "0"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   350
      Left            =   2040
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtEDX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtECX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtEBX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtEAX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblLevel 
      Caption         =   "Level"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblEBX 
      Caption         =   "EBX"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblEAX 
      Caption         =   "EAX"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblEDX 
      Caption         =   "EDX"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblECX 
      Caption         =   "ECX"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmCPUIDOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmCPUIDOther"


Private Sub cmdGetData_Click()
On Error GoTo VB_Error
    
    txtLevel.Text = MinMax(Val(txtLevel.Text), 0, 4294967295#)
    
    
    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    Call cpuid_(uint32_int32(txtLevel.Text), outEAX, outEBX, outECX, outEDX)
    
    txtEAX.Text = Right$("00000000" & ltoa_(outEAX, 16), 8)
    txtEBX.Text = Right$("00000000" & ltoa_(outEBX, 16), 8)
    txtECX.Text = Right$("00000000" & ltoa_(outECX, 16), 8)
    txtEDX.Text = Right$("00000000" & ltoa_(outEDX, 16), 8)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdGetData_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    txtLevel.Text = uint32_int32(Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\CPUIDOther", "Level")))
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    txtLevel.Text = MinMax(Val(txtLevel.Text), 0, 4294967295#)
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\CPUIDOther", "Level", uint32_int32(txtLevel.Text), REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
