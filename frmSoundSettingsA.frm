VERSION 5.00
Begin VB.Form frmSoundSettingsA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound Settings Accessibility"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmSoundSettingsA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox chkScreenReader 
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkShowSounds 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblScreenReader 
      Caption         =   "Screen Reader"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblShowSounds 
      Caption         =   "Show Sounds"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmSoundSettingsA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmSoundSettingsA"


Private Sub cmdApply_Click()
On Error GoTo VB_Error

    If WinVersion(0, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETSCREENREADER, CBool(chkScreenReader.Value), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    'If SystemParametersInfo(SPI_SETSHOWSOUNDS, CBool(chkShowSounds.Value), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    Dim bValue As Boolean
    
    If WinVersion(0, 5000000, True) = True Then
        If SystemParametersInfo(SPI_GETSCREENREADER, 0&, bValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkScreenReader.Value = IIf(bValue, 1, 0)
    Else
        lblScreenReader.Enabled = False
        chkScreenReader.Enabled = False
    End If

    If SystemParametersInfo(SPI_GETSHOWSOUNDS, 0&, bValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    chkShowSounds.Value = IIf(bValue, 1, 0)

Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub
