VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDisplaySettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Settings"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmDisplaySettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtBitsPerPixel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CheckBox chkAllModes 
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox chkTest 
      Height          =   255
      Left            =   5520
      TabIndex        =   16
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox chkGlobal 
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   4800
      TabIndex        =   18
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   3720
      TabIndex        =   17
      Top             =   4200
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwModes 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblVRefresh 
      Caption         =   "VRefresh"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblHeight 
      Caption         =   "Height"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblWidth 
      Caption         =   "Width"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblBitsPerPixel 
      Caption         =   "BPP"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblCurrentMode 
      Caption         =   "Current Mode"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblAllModes 
      Caption         =   "All Modes"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblTest 
      Caption         =   "Test"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblGlobal 
      Caption         =   "Global Change"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Modes Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDisplaySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmDisplaySettings"


Private Sub chkAllModes_Click()
On Error GoTo VB_Error

    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkAllModes_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    If lvwModes.SelectedItem Is Nothing Then Exit Sub
    
    
    Dim DEVMODE As DEVMODE
    With DEVMODE
        .dmSize = Len(DEVMODE)
        .dmBitsPerPel = lvwModes.SelectedItem.SubItems(2)
        .dmPelsWidth = lvwModes.SelectedItem.Text
        .dmPelsHeight = lvwModes.SelectedItem.SubItems(1)
        .dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_DISPLAYFREQUENCY
        .dmDisplayFrequency = lvwModes.SelectedItem.SubItems(3)
        '.dmPosition 'Multimonitor
    End With
    
    
    If chkTest.value = 1 Then
        If ChangeDisplaySettings(DEVMODE, CDS_TEST) <> 0 Then
            If MessageBoxEx(0&, "Display test failed. Mode was not set.", "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "MessageBoxEx")
            Exit Sub
        End If
    End If
    If chkGlobal.value = 1 Then
        lErrors = ChangeDisplaySettings(DEVMODE, CDS_UPDATEREGISTRY Or CDS_GLOBAL)
    Else
        lErrors = ChangeDisplaySettings(DEVMODE, CDS_UPDATEREGISTRY)
    End If
    
    Select Case lErrors
        Case DISP_CHANGE_RESTART: If MessageBoxEx(0&, "Must restart Windows for changes to be implemented.", "Restart", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "MessageBoxEx")
        Case DISP_CHANGE_BADFLAGS: Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "ChangeDisplaySettings")
        Case DISP_CHANGE_BADPARAM: Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "ChangeDisplaySettings")
        Case DISP_CHANGE_FAILED: Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "ChangeDisplaySettings")
        Case DISP_CHANGE_BADMODE: Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "ChangeDisplaySettings")
        Case DISP_CHANGE_NOTUPDATED: Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "ChangeDisplaySettings")
        Case DISP_CHANGE_BADDUALVIEW: Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "ChangeDisplaySettings")
    End Select
    
    
    txtBitsPerPixel.Text = GetDeviceCaps(frmDisplaySettings.hdc, BITSPIXEL)
    txtWidth.Text = Screen.Width \ Screen.TwipsPerPixelX
    txtHeight.Text = Screen.Height \ Screen.TwipsPerPixelY
    If WinVersion(-1, 0, True) = True Then
        txtVRefresh.Text = GetDeviceCaps(frmDisplaySettings.hdc, VREFRESH)
    Else
        txtVRefresh.Text = Reg_Read(HKEY_CURRENT_CONFIG, "Display\Settings", "RefreshRate")
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    Call ListView_Clear(lvwModes)
    
    
    txtBitsPerPixel.Text = GetDeviceCaps(frmDisplaySettings.hdc, BITSPIXEL)
    txtWidth.Text = Screen.Width \ Screen.TwipsPerPixelX
    txtHeight.Text = Screen.Height \ Screen.TwipsPerPixelY
    If WinVersion(-1, 0, True) = True Then
        txtVRefresh.Text = GetDeviceCaps(frmDisplaySettings.hdc, VREFRESH)
    Else
        txtVRefresh.Text = Reg_Read(HKEY_CURRENT_CONFIG, "Display\Settings", "RefreshRate")
    End If
    
    
    Dim DEVMODE As DEVMODE
    Dim lIncrement As Long
    Dim bExtended As Boolean
    Dim lFlags As Long
    
    bExtended = WinVersion(4010000, 5000000, True)
    If chkAllModes.value = 1 Then lFlags = EDS_RAWMODE
    DEVMODE.dmSize = Len(DEVMODE)
    
    Do
        If bExtended = True Then
            If EnumDisplaySettingsEx(ByVal 0&, lIncrement, DEVMODE, lFlags) = 0 Then
                If Err.LastDllError <> 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "EnumDisplaySettings")
                Exit Do
            End If
        Else
            If EnumDisplaySettings(ByVal 0&, lIncrement, DEVMODE) = 0 Then
                If Err.LastDllError <> 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "EnumDisplaySettings")
                Exit Do
            End If
        End If
        
        lIncrement = lIncrement + 1
        
        With DEVMODE
            With lvwModes.ListItems.Add(, , DEVMODE.dmPelsWidth)
                .SubItems(1) = DEVMODE.dmPelsHeight
                .SubItems(2) = DEVMODE.dmBitsPerPel
                .SubItems(3) = DEVMODE.dmDisplayFrequency
            End With
            
            'If iBPP = .dmBitsPerPel And _
            '    iWidth = .dmPelsWidth And _
            '    iHeight = .dmPelsHeight And _
            '    iVRefresh = .dmDisplayFrequency Then
            '    lvwModes.ListItems.Item(lvwModes.ListItems.Count).Selected = True
            'End If
        End With
        
        If bShutdown = True Then Exit Do
    Loop
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdRefresh_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwModes.ColumnHeaders
        .Add , , "Width-Pixel"
        .Add , , "Height-Pixel"
        .Add , , "Bits Per Pixel"
        .Add , , "Verticle Refresh-HZ"
    End With
    
    If WinVersion(4010000, 5000000, True) = False Then chkAllModes.Enabled = False
    chkGlobal.value = IIf(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "GlobalChange"), 1, 0)
    chkTest.value = IIf(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "Test"), 1, 0)
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "GlobalChange", chkGlobal.value, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "Test", chkTest.value, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
