VERSION 5.00
Begin VB.Form frmWindowsSettingsA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Settings Accessibility"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   Icon            =   "frmWindowsSettingsA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFocusBorderWidth 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtFocusBorderHeight 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtPIX1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Text            =   "PIX"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtPIX2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "PIX"
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblFocusBorderWidth 
      Caption         =   "Focus Border Width"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblFocusBorderHeight 
      Caption         =   "Focus Border Height"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmWindowsSettingsA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmWindowsSettings"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtFocusBorderHeight.Text = MinMax(Val(txtFocusBorderHeight.Text), 0, 2147483647)
    txtFocusBorderWidth.Text = MinMax(Val(txtFocusBorderWidth.Text), 0, 2147483647)
    
    
    If SystemParametersInfo(SPI_SETFOCUSBORDERHEIGHT, 0&, ByVal CLng(txtFocusBorderHeight.Text), SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    If SystemParametersInfo(SPI_SETFOCUSBORDERWIDTH, 0&, ByVal CLng(txtFocusBorderWidth.Text), SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If WinVersion(-1, 5010000, True) = True Then
        Dim lValue As Long
        
        If SystemParametersInfo(SPI_GETFOCUSBORDERHEIGHT, 0&, lValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtFocusBorderHeight.Text = lValue
        If SystemParametersInfo(SPI_GETFOCUSBORDERWIDTH, 0&, lValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtFocusBorderWidth.Text = lValue
    Else
        lblFocusBorderHeight.Enabled = False
        txtFocusBorderHeight.Enabled = False
        lblFocusBorderWidth.Enabled = False
        txtFocusBorderWidth.Enabled = False
        cmdApply.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
