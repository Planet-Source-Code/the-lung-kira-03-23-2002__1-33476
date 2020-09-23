VERSION 5.00
Begin VB.Form frmMouseSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Settings"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmMouseSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSpeed 
      Height          =   285
      Left            =   1920
      TabIndex        =   29
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtCursorTrails 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtDataQueueSize 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtMS2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "MS"
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtMS1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "MS"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtPIX4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   3360
      Width           =   255
   End
   Begin VB.TextBox txtPIX3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox txtPIX2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtPIX1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkHotTracking 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkCursorShadow 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtWheelScrollLines 
      Height          =   285
      Left            =   1920
      TabIndex        =   33
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CheckBox chkSnapToDefault 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkSwapButton 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox txtHoverTimeHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   21
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtHoverTimeWidth 
      Height          =   285
      Left            =   1920
      TabIndex        =   24
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtDoubleClickWidth 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtDoubleClickHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtHoverTime 
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2520
      TabIndex        =   34
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtDoubleClickTime 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblDataQueueSize 
      Caption         =   "Data Queue Size"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblHotTracking 
      Caption         =   "Hot Tracking"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblCursorTrails 
      Caption         =   "Cursor Trails"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblCursorShadow 
      Caption         =   "Cursor Shadow"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblWheelScrollLines 
      Caption         =   "Wheel Scroll Lines"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblSnapToDefault 
      Caption         =   "Snap To Default"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblSwapButton 
      Caption         =   "Swap Button"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblHoverTimeHeight 
      Caption         =   "Hover Time Height"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblHoverTimeWidth 
      Caption         =   "Hover Time Width"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblDoubleClickWidth 
      Caption         =   "Double Click Width"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblDoubleClickHeight 
      Caption         =   "Double Click Height"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblHoverTime 
      Caption         =   "Hover Time"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblDoubleClickTime 
      Caption         =   "Double Click Time"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmMouseSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMouseSettings"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtCursorTrails.Text = MinMax(Val(txtCursorTrails.Text), 0, 16)
    txtDataQueueSize.Text = MinMax(Val(txtDataQueueSize.Text), 0, 4294967295#)
    txtDoubleClickHeight.Text = MinMax(Val(txtDoubleClickHeight.Text), 0, 2147483647)
    txtDoubleClickTime.Text = MinMax(Val(txtDoubleClickTime.Text), 0, 5000)
    txtDoubleClickWidth.Text = MinMax(Val(txtDoubleClickWidth.Text), 0, 2147483647)
    txtHoverTime.Text = MinMax(Val(txtHoverTime.Text), 0, 2147483647)
    txtHoverTimeHeight.Text = MinMax(Val(txtHoverTimeHeight.Text), 0, 2147483647)
    txtHoverTimeWidth.Text = MinMax(Val(txtHoverTimeWidth.Text), 0, 2147483647)
    txtSpeed.Text = MinMax(Val(txtSpeed.Text), 1, 20)
    txtWheelScrollLines.Text = MinMax(Val(txtWheelScrollLines.Text), 0, 2147483647)
    
    
    If SetDoubleClickTime(txtDoubleClickTime.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetDoubleClickTime")
    
    
    If WinVersion(-1, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETCURSORSHADOW, 0&, ByVal CBool(chkCursorShadow.value), SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    If WinVersion(0, 5010000, True) = True Then
        If SystemParametersInfo(SPI_SETMOUSETRAILS, CLng(txtCursorTrails.Text), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    If WinVersion(-1, 0, True) = True Then
        If lblDataQueueSize.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Mouclass\Parameters", "MouseDataQueueSize", uint32_int32(txtDataQueueSize.Text), REG_DWORD)
    End If
    
    If SystemParametersInfo(SPI_SETDOUBLECLKHEIGHT, CLng(txtDoubleClickHeight.Text), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    If SystemParametersInfo(SPI_SETDOUBLECLKWIDTH, CLng(txtDoubleClickWidth.Text), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
    If WinVersion(4010000, 0, True) = True Then
        If SystemParametersInfo(SPI_SETMOUSEHOVERTIME, CLng(txtHoverTime.Text), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        If SystemParametersInfo(SPI_SETMOUSEHOVERHEIGHT, CLng(txtHoverTimeHeight.Text), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        If SystemParametersInfo(SPI_SETMOUSEHOVERWIDTH, CLng(txtHoverTimeWidth.Text), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        
        If SystemParametersInfo(SPI_SETSNAPTODEFBUTTON, CBool(chkSnapToDefault.value), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        
        If SystemParametersInfo(SPI_SETWHEELSCROLLLINES, CLng(txtWheelScrollLines.Text), 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETHOTTRACKING, 0&, ByVal CBool(chkHotTracking.value), SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        If SystemParametersInfo(SPI_SETMOUSESPEED, 0&, ByVal CLng(txtSpeed.Text), SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    
    If SystemParametersInfo(SPI_SETMOUSEBUTTONSWAP, chkSwapButton.value, 0&, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    Dim bValue As Byte
    Dim lValue As Long
    Dim bFail As Byte
    
    If WinVersion(-1, 5000000, True) = True Then
        If SystemParametersInfo(SPI_GETCURSORSHADOW, 0&, bValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkCursorShadow.value = IIf(bValue, 1, 0)
    Else
        lblCursorShadow.Enabled = False
        chkCursorShadow.Enabled = False
    End If
    If WinVersion(0, 5010000, True) = True Then
        If SystemParametersInfo(SPI_GETMOUSETRAILS, 0&, lValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtCursorTrails.Text = CStr(lValue)
    Else
        lblCursorTrails.Enabled = False
        txtCursorTrails.Enabled = False
    End If
    If WinVersion(-1, 0, True) = True Then
        lValue = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Mouclass\Parameters", "MouseDataQueueSize", bFail)
        If bFail <> 0 Then
            lblDataQueueSize.Enabled = False
        Else
            txtDataQueueSize.Text = int32_uint32(lValue)
        End If
    End If
    
    txtDoubleClickTime.Text = CStr(GetDoubleClickTime())
    txtDoubleClickHeight.Text = CStr(GetSystemMetrics(SM_CYDOUBLECLK))
    txtDoubleClickWidth.Text = CStr(GetSystemMetrics(SM_CXDOUBLECLK))
    
    If WinVersion(4010000, 0, True) = True Then
        If SystemParametersInfo(SPI_GETMOUSEHOVERTIME, 0&, lValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtHoverTime.Text = CStr(lValue)
        If SystemParametersInfo(SPI_GETMOUSEHOVERHEIGHT, 0&, lValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtHoverTimeHeight.Text = CStr(lValue)
        If SystemParametersInfo(SPI_GETMOUSEHOVERWIDTH, 0&, lValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtHoverTimeWidth.Text = CStr(lValue)
        
        If SystemParametersInfo(SPI_GETSNAPTODEFBUTTON, 0&, bValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkSnapToDefault.value = IIf(bValue, 1, 0)
        
        If SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0&, lValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtWheelScrollLines.Text = CStr(lValue)
    Else
        lblHoverTime.Enabled = False
        txtHoverTime.Enabled = False
        lblHoverTimeHeight.Enabled = False
        txtHoverTimeHeight.Enabled = False
        lblHoverTimeWidth.Enabled = False
        txtHoverTimeWidth.Enabled = False
        lblSnapToDefault.Enabled = False
        chkSnapToDefault.Enabled = False
        lblWheelScrollLines.Enabled = False
        txtWheelScrollLines.Enabled = False
    End If
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_GETHOTTRACKING, 0&, bValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkHotTracking.value = IIf(bValue, 1, 0)
        
        If SystemParametersInfo(SPI_GETMOUSESPEED, 0&, lValue, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtSpeed.Text = CStr(lValue)
    Else
        lblSpeed.Enabled = False
        txtSpeed.Enabled = False
    End If
    
    chkSwapButton.value = GetSystemMetrics(SM_SWAPBUTTON)
    
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
