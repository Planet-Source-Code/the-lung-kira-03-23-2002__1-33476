VERSION 5.00
Begin VB.Form frmMouseKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Keys"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmMouseKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "MS"
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtCtrlSpeed 
      Height          =   285
      Left            =   5160
      TabIndex        =   27
      Text            =   "0"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtTimeToMaxSpeed 
      Height          =   285
      Left            =   4800
      TabIndex        =   31
      Text            =   "0"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtMaxSpeed 
      Height          =   285
      Left            =   5160
      TabIndex        =   29
      Text            =   "0"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox chkLeftButtonDown 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkRightButtonDown 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6000
      TabIndex        =   25
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkRightButtonSelect 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkLeftButtonSelect 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkReplaceNumbers 
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkMouseMode 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkModifiers 
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkMouseKeysOn 
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkIndicator 
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeySound 
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeyActive 
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chkConfirmHotKey 
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox chkAvailable 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5280
      TabIndex        =   33
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblCtrlSpeed 
      Caption         =   "Ctrl Speed"
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblTimeToMaxSpeed 
      Caption         =   "Time To Max Speed"
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblMaxSpeed 
      Caption         =   "Max Speed"
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblLeftButtonDown 
      Caption         =   "Left Button Down"
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblRightButtonDown 
      Caption         =   "Right Button Down"
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblRightButtonSelect 
      Caption         =   "Right Button Select"
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblLeftButtonSelect 
      Caption         =   "Left Button Select"
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblReplaceNumbers 
      Caption         =   "Replace Numbers"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblMouseMode 
      Caption         =   "Mouse Mode"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblModifiers 
      Caption         =   "Modifiers"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblMouseKeysOn 
      Caption         =   "Mouse Keys On"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblIndicator 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblHotKeySound 
      Caption         =   "Hot Key Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblHotKeyActive 
      Caption         =   "Hot Key Active"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblConfirmHotKey 
      Caption         =   "Confirm Hot Key"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMouseKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMouseKeys"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    If WinVersion(-1, 0, True) = True Then
        txtCtrlSpeed.Text = MinMax(Val(txtCtrlSpeed.Text), 10, 360)
    Else
        txtCtrlSpeed.Text = MinMax(Val(txtCtrlSpeed.Text), 0, 2147483647)
    End If
    txtMaxSpeed.Text = MinMax(Val(txtMaxSpeed.Text), 0, 2147483647)
    txtTimeToMaxSpeed.Text = MinMax(Val(txtTimeToMaxSpeed.Text), 1000, 5000)
    
    
    Dim MOUSEKEYS As MOUSEKEYS
    With MOUSEKEYS
        .cbSize = Len(MOUSEKEYS)
        
        .dwFlags = .dwFlags Or MKF_AVAILABLE
        
        If WinVersion(4000000, 5000000, True) = True Then
            If chkConfirmHotKey.value = 1 Then .dwFlags = .dwFlags Or MKF_CONFIRMHOTKEY
            If chkIndicator.value = 1 Then .dwFlags = .dwFlags Or MKF_INDICATOR
            If chkModifiers.value = 1 Then .dwFlags = .dwFlags Or MKF_MODIFIERS
            If chkReplaceNumbers.value = 1 Then .dwFlags = .dwFlags Or MKF_REPLACENUMBERS
        End If
        
        If chkHotKeyActive.value = 1 Then .dwFlags = .dwFlags Or MKF_HOTKEYACTIVE
        If chkHotKeySound.value = 1 Then .dwFlags = .dwFlags Or MKF_HOTKEYSOUND
        If chkMouseKeysOn.value = 1 Then .dwFlags = .dwFlags Or MKF_MOUSEKEYSON
        
        .iCtrlSpeed = txtCtrlSpeed.Text
        .iMaxSpeed = txtMaxSpeed.Text
        .iTimeToMaxSpeed = txtTimeToMaxSpeed.Text
    End With
    
    If SystemParametersInfo(SPI_SETMOUSEKEYS, MOUSEKEYS.cbSize, MOUSEKEYS, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim MOUSEKEYS As MOUSEKEYS
    MOUSEKEYS.cbSize = Len(MOUSEKEYS)
    
    If SystemParametersInfo(SPI_GETMOUSEKEYS, MOUSEKEYS.cbSize, MOUSEKEYS, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    
    If MOUSEKEYS.dwFlags And MKF_AVAILABLE Then
        With MOUSEKEYS
            If .dwFlags And MKF_AVAILABLE Then chkAvailable.value = 1
            If .dwFlags And MKF_HOTKEYACTIVE Then chkHotKeyActive.value = 1
            If .dwFlags And MKF_HOTKEYSOUND Then chkHotKeySound.value = 1
            
            If WinVersion(4000000, 5000000, True) = True Then
                If .dwFlags And MKF_CONFIRMHOTKEY Then chkConfirmHotKey.value = 1
                If .dwFlags And MKF_INDICATOR Then chkIndicator.value = 1
                If .dwFlags And MKF_MODIFIERS Then chkModifiers.value = 1
                If .dwFlags And MKF_REPLACENUMBERS Then chkReplaceNumbers.value = 1
            Else
                lblConfirmHotKey.Enabled = False
                chkConfirmHotKey.Enabled = False
                lblIndicator.Enabled = False
                chkIndicator.Enabled = False
                lblModifiers.Enabled = False
                chkModifiers.Enabled = False
                lblReplaceNumbers.Enabled = False
                chkReplaceNumbers.Enabled = False
            End If
            If WinVersion(4010000, 5000000, True) = True Then
                If .dwFlags And MKF_MOUSEMODE Then chkMouseMode.value = 1
                If .dwFlags And MKF_LEFTBUTTONSEL Then chkLeftButtonSelect.value = 1
                If .dwFlags And MKF_RIGHTBUTTONSEL Then chkRightButtonSelect.value = 1
                If .dwFlags And MKF_LEFTBUTTONDOWN Then chkLeftButtonDown.value = 1
                If .dwFlags And MKF_RIGHTBUTTONDOWN Then chkRightButtonDown.value = 1
            Else
                lblMouseMode.Enabled = False
                lblLeftButtonSelect.Enabled = False
                lblRightButtonSelect.Enabled = False
                lblLeftButtonDown.Enabled = False
                lblRightButtonDown.Enabled = False
            End If
            
            txtCtrlSpeed.Text = .iCtrlSpeed
            txtMaxSpeed.Text = .iMaxSpeed
            txtTimeToMaxSpeed.Text = .iTimeToMaxSpeed
        End With
    Else
        lblHotKeyActive.Enabled = False
        chkHotKeyActive.Enabled = False
        lblHotKeySound.Enabled = False
        chkHotKeySound.Enabled = False
        lblMouseKeysOn.Enabled = False
        chkMouseKeysOn.Enabled = False
        
        lblConfirmHotKey.Enabled = False
        chkConfirmHotKey.Enabled = False
        lblIndicator.Enabled = False
        chkIndicator.Enabled = False
        lblModifiers.Enabled = False
        chkModifiers.Enabled = False
        lblReplaceNumbers.Enabled = False
        chkReplaceNumbers.Enabled = False
                
        lblMouseMode.Enabled = False
        lblLeftButtonSelect.Enabled = False
        lblRightButtonSelect.Enabled = False
        lblLeftButtonDown.Enabled = False
        lblRightButtonDown.Enabled = False
                
        cmdApply.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
