VERSION 5.00
Begin VB.Form frmFilterKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Keys"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmFilterKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMS4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      TabIndex        =   25
      Text            =   "MS"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtMS3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      TabIndex        =   22
      Text            =   "MS"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtMS2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      TabIndex        =   19
      Text            =   "MS"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtMS1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      TabIndex        =   16
      Text            =   "MS"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtBounce 
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Text            =   "0"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtRepeat 
      Height          =   285
      Left            =   3360
      TabIndex        =   21
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtDelay 
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Text            =   "0"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtWait 
      Height          =   285
      Left            =   3360
      TabIndex        =   24
      Text            =   "0"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkIndicator 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkConfirmHotKey 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeySound 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeyActive 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox chkFilterKeysOn 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkClickOn 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3960
      TabIndex        =   26
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox chkAvailable 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblBounce 
      Caption         =   "Bounce"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblRepeat 
      Caption         =   "Repeat"
      Height          =   255
      Left            =   2280
      TabIndex        =   20
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblDelay 
      Caption         =   "Delay"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblWait 
      Caption         =   "Wait"
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblIndicator 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblConfirmHotKey 
      Caption         =   "Confirm Hot Key"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblHotKeySound 
      Caption         =   "Hot Key Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblHotKeyActive 
      Caption         =   "Hot Key Active"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblFilterKeysOn 
      Caption         =   "Filter Keys On"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblClickOn 
      Caption         =   "Click On"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmFilterKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmFilterKeys"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtBounce.Text = MinMax(Val(txtBounce.Text), 0, 2147483647)
    txtDelay.Text = MinMax(Val(txtDelay.Text), 0, 2147483647)
    txtRepeat.Text = MinMax(Val(txtRepeat.Text), 0, 2147483647)
    txtWait.Text = MinMax(Val(txtWait.Text), 0, 2147483647)
    
    
    Dim FILTERKEYS As FILTERKEYS
    With FILTERKEYS
        .cbSize = Len(FILTERKEYS)
        
        .dwFlags = .dwFlags Or FKF_AVAILABLE
        If WinVersion(4000000, 5000000, True) = True Then
            If chkConfirmHotKey.value = 1 Then .dwFlags = .dwFlags Or FKF_CONFIRMHOTKEY
            If chkIndicator.value = 1 Then .dwFlags = .dwFlags Or FKF_INDICATOR
        End If
        If chkClickOn.value = 1 Then .dwFlags = .dwFlags Or FKF_CLICKON
        If chkFilterKeysOn.value = 1 Then .dwFlags = .dwFlags Or FKF_FILTERKEYSON
        If chkHotKeyActive.value = 1 Then .dwFlags = .dwFlags Or FKF_HOTKEYACTIVE
        If chkHotKeySound.value = 1 Then .dwFlags = .dwFlags Or FKF_HOTKEYSOUND
        
        .iBounceMSec = txtBounce.Text
        .iDelayMSec = txtDelay.Text
        .iRepeatMSec = txtRepeat.Text
        .iWaitMSec = txtWait.Text
    End With
    
    If SystemParametersInfo(SPI_SETFILTERKEYS, FILTERKEYS.cbSize, FILTERKEYS, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim FILTERKEYS As FILTERKEYS
    FILTERKEYS.cbSize = Len(FILTERKEYS)
    
    If SystemParametersInfo(SPI_GETFILTERKEYS, FILTERKEYS.cbSize, FILTERKEYS, 0) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    
    If FILTERKEYS.dwFlags And FKF_AVAILABLE Then
        With FILTERKEYS
            If .dwFlags And FKF_AVAILABLE Then chkAvailable.value = 1
            If .dwFlags And FKF_CLICKON Then chkClickOn.value = 1
            If .dwFlags And FKF_FILTERKEYSON Then chkFilterKeysOn.value = 1
            If .dwFlags And FKF_HOTKEYACTIVE Then chkHotKeyActive.value = 1
            If .dwFlags And FKF_HOTKEYSOUND Then chkHotKeySound.value = 1
            
            If WinVersion(4000000, 5000000, True) = True Then
                If .dwFlags And FKF_CONFIRMHOTKEY Then chkConfirmHotKey.value = 1
                If .dwFlags And FKF_INDICATOR Then chkIndicator.value = 1
            Else
                chkConfirmHotKey.Enabled = False
                chkIndicator.Enabled = False
            End If
            
            txtWait.Text = .iWaitMSec
            txtDelay.Text = .iDelayMSec
            txtRepeat.Text = .iRepeatMSec
            txtBounce.Text = .iBounceMSec
        End With
    Else
        lblClickOn.Enabled = False
        chkClickOn.Enabled = False
        lblFilterKeysOn.Enabled = False
        chkFilterKeysOn.Enabled = False
        lblHotKeyActive.Enabled = False
        chkHotKeyActive.Enabled = False
        lblHotKeySound.Enabled = False
        chkHotKeySound.Enabled = False
        lblConfirmHotKey.Enabled = False
        chkConfirmHotKey.Enabled = False
        lblIndicator.Enabled = False
        chkIndicator.Enabled = False
        lblWait.Enabled = False
        txtWait.Enabled = False
        lblDelay.Enabled = False
        txtDelay.Enabled = False
        lblRepeat.Enabled = False
        txtRepeat.Enabled = False
        lblBounce.Enabled = False
        txtBounce.Enabled = False
        cmdApply.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
