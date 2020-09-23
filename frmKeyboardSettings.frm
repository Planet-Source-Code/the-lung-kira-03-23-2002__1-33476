VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKeyboardSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Settings"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmKeyboardSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDataQueueSize 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtMS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "MS"
      Top             =   120
      Width           =   255
   End
   Begin MSComctlLib.Slider sldRepeatRate 
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   3240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Max             =   31
   End
   Begin VB.ComboBox cboLanguageToggle 
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CheckBox chkCues 
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2880
      TabIndex        =   19
      Top             =   3600
      Width           =   975
   End
   Begin VB.CheckBox chkPref 
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtBlinkRate 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.Slider sldRepeatDelay 
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   2400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Max             =   3
   End
   Begin VB.Label lblDataQueueSize 
      Caption         =   "Data Queue Size"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblLanguageToggle 
      Caption         =   "Language Toggle"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblCues 
      Caption         =   "Cues Underlined"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblPref 
      Caption         =   "Keyboard Preference"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblFast 
      Caption         =   "Fast"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblSlow 
      Caption         =   "Slow"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblLong 
      Caption         =   "Long"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblShort 
      Caption         =   "Short"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblRepeatRate 
      Caption         =   "Repeat Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblRepeatDelay 
      Caption         =   "Repeat Delay"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblBlinkRate 
      Caption         =   "Caret Blink Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmKeyboardSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmKeyboardSettings"


Private Sub cboLanguageToggle_Click()
On Error GoTo VB_Error

    If lblLanguageToggle.Enabled = False Then lblLanguageToggle.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboLanguageToggle_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtBlinkRate.Text = MinMax(Val(txtBlinkRate.Text), 0, 5000)
    txtDataQueueSize.Text = MinMax(Val(txtDataQueueSize.Text), 0, 4294967295#)
    
    
    If SetCaretBlinkTime(CLng(txtBlinkRate.Text)) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetCaretBlinkTime")
    
    If WinVersion(-1, 0, True) = True Then
        If lblDataQueueSize.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Kbdclass\Parameters", "KeyboardDataQueueSize", uint32_int32(txtDataQueueSize.Text), REG_DWORD)
    End If
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETKEYBOARDCUES, 0&, ByVal CBool(chkCues.value), SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    If WinVersion(0, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETKEYBOARDPREF, CBool(chkPref.value), 0&, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    
    If SystemParametersInfo(SPI_SETKEYBOARDDELAY, CLng(sldRepeatDelay.value), 0&, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    If SystemParametersInfo(SPI_SETKEYBOARDSPEED, CLng(sldRepeatRate.value), 0&, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
    If cboLanguageToggle.ListIndex > -1 Then
        If lblLanguageToggle.Enabled = True Then Call Reg_Write(HKEY_CURRENT_USER, "Keyboard Layout\Toggle", "Hotkey", (cboLanguageToggle.ListIndex + 1), REG_SZ)
    End If
    If SystemParametersInfo(SPI_SETLANGTOGGLE, 0&, 0&, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With cboLanguageToggle
        .AddItem "ALT+SHIFT"
        .AddItem "CTRL+SHIFT"
        .AddItem "None"
    End With
    
    
    Dim bFail As Byte
    Dim bValue As Boolean
    Dim iValue As Integer
    Dim lValue As Long
    Dim lReturn As Long
    
    txtBlinkRate.Text = GetCaretBlinkTime()
    
    If WinVersion(-1, 0, True) = True Then
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Kbdclass\Parameters", "KeyboardDataQueueSize", bFail)
        If bFail <> 0 Then
            lblDataQueueSize.Enabled = False
        Else
            txtDataQueueSize.Text = int32_uint32(lReturn)
        End If
    Else
        lblDataQueueSize.Enabled = False
        txtDataQueueSize.Enabled = False
    End If
    
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_GETKEYBOARDCUES, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkCues.value = IIf(bValue, 1, 0)
    Else
        lblCues.Enabled = False
        chkCues.Enabled = False
    End If
    
    
    If SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0&, lValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    sldRepeatDelay.value = MinMax(lValue, 0, 3)
    
    
    If WinVersion(0, 5000000, True) = True Then
        Dim bPref As Boolean
        If SystemParametersInfo(SPI_GETKEYBOARDPREF, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkPref.value = IIf(bValue, 1, 0)
    Else
        lblPref.Enabled = False
        chkPref.Enabled = False
    End If
    
    
    If SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0&, lValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    sldRepeatRate.value = MinMax(lValue, 0, 31)
    
    
    lReturn = Reg_Read(HKEY_CURRENT_USER, "Keyboard Layout\Toggle", "Hotkey", bFail)
    If bFail <> 0 Then
        lblLanguageToggle.Enabled = False
    Else
        Select Case lReturn
            Case "1": cboLanguageToggle.ListIndex = 0
            Case "2": cboLanguageToggle.ListIndex = 1
            Case "3": cboLanguageToggle.ListIndex = 2
            Case Else: cboLanguageToggle.ListIndex = -1
        End Select
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub txtDataQueueSize_Change()
On Error GoTo VB_Error

    If lblDataQueueSize.Enabled = False Then lblDataQueueSize.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtDataQueueSize_Change")
Resume Next
End Sub
