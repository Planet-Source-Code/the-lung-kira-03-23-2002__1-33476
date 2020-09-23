VERSION 5.00
Begin VB.Form frmAccessTimeout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access Timeout"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmAccessTimeout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Text            =   "MS"
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox chkTimeOutOn 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkOnOffFeedback 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "0"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblTimeOutOn 
      Caption         =   "Time Out On"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblOnOffFeedback 
      Caption         =   "On Off Feedback"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblTimeOut 
      Caption         =   "Time Out"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmAccessTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmAccessTimeout"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtTimeOut.Text = MinMax(Val(txtTimeOut.Text), 0, 2147483647)
    
    
    Dim ACCESSTIMEOUT As ACCESSTIMEOUT
    With ACCESSTIMEOUT
        .cbSize = Len(ACCESSTIMEOUT)
        
        If chkOnOffFeedback.value = 1 Then .dwFlags = .dwFlags Or ATF_ONOFFFEEDBACK
        If chkTimeOutOn.value = 1 Then .dwFlags = .dwFlags Or ATF_TIMEOUTON
        
        .iTimeOutMSec = txtTimeOut.Text
    End With
    
    If SystemParametersInfo(SPI_SETACCESSTIMEOUT, ACCESSTIMEOUT.cbSize, ACCESSTIMEOUT, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim ACCESSTIMEOUT As ACCESSTIMEOUT
    ACCESSTIMEOUT.cbSize = Len(ACCESSTIMEOUT)
    
    If SystemParametersInfo(SPI_GETACCESSTIMEOUT, ACCESSTIMEOUT.cbSize, ACCESSTIMEOUT, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    
    With ACCESSTIMEOUT
        If .dwFlags And ATF_ONOFFFEEDBACK Then chkOnOffFeedback.value = 1
        If .dwFlags And ATF_TIMEOUTON Then chkTimeOutOn.value = 1
        
        txtTimeOut.Text = .iTimeOutMSec
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
