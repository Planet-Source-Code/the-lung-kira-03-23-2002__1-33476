VERSION 5.00
Begin VB.Form frmMouseSettingsA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Settings Accessibility"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmMouseSettingsA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "MS"
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox chkMouseVanish 
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chkMouseSonar 
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtClickLockTime 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox chkClickLock 
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblMouseVanish 
      Caption         =   "Mouse Vanish"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblMouseSonar 
      Caption         =   "Mouse Sonar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblClickLockTime 
      Caption         =   "Click Lock Time"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblClickLock 
      Caption         =   "Click Lock"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMouseSettingsA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMouseSettingsA"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtClickLockTime.Text = MinMax(Val(txtClickLockTime.Text), 0, 2147483647)
    
    
    If SystemParametersInfo(SPI_SETMOUSECLICKLOCK, 0&, ByVal CBool(chkClickLock.value), SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    If SystemParametersInfo(SPI_SETMOUSECLICKLOCKTIME, 0&, ByVal CLng(txtClickLockTime.Text), SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    If SystemParametersInfo(SPI_SETMOUSESONAR, 0&, ByVal CBool(chkMouseSonar.value), SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    If SystemParametersInfo(SPI_SETMOUSEVANISH, 0&, ByVal CBool(chkMouseVanish.value), SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If WinVersion(4900000, 5010000, True) = True Then
        Dim bValue As Boolean
        Dim lValue As Long
        
        If SystemParametersInfo(SPI_GETMOUSECLICKLOCK, 0&, bValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkClickLock.value = IIf(bValue, 1, 0)
        If SystemParametersInfo(SPI_GETMOUSECLICKLOCKTIME, 0&, lValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtClickLockTime.Text = lValue
        If SystemParametersInfo(SPI_GETMOUSESONAR, 0&, bValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkMouseSonar.value = IIf(bValue, 1, 0)
        If SystemParametersInfo(SPI_GETMOUSEVANISH, 0&, bValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkMouseVanish.value = IIf(bValue, 1, 0)
    Else
        lblClickLock.Enabled = False
        chkClickLock.Enabled = False
        lblClickLockTime.Enabled = False
        txtClickLockTime.Enabled = False
        lblMouseSonar.Enabled = False
        chkMouseSonar.Enabled = False
        lblMouseVanish.Enabled = False
        chkMouseVanish.Enabled = False
        cmdApply.Enabled = False
    End If

Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
