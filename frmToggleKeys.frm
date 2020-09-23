VERSION 5.00
Begin VB.Form frmToggleKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toggle Keys"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   Icon            =   "frmToggleKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   1440
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox chkAvailable 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkToggleKeysOn 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeyActive 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeySound 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkConfirmHotKey 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblToggleKeysOn 
      Caption         =   "Toggle Keys On"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblHotKeyActive 
      Caption         =   "Hot Key Active"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblHotKeySound 
      Caption         =   "Hot Key Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblConfirmHotKey 
      Caption         =   "Confirm Hot Key"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmToggleKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmToggleKeys"


Private Sub cmdApply_Click()
On Error GoTo VB_Error

    Dim TOGGLEKEYS As TOGGLEKEYS
    With TOGGLEKEYS
        .cbSize = Len(TOGGLEKEYS)
        
        .dwFlags = .dwFlags Or TKF_AVAILABLE
        If WinVersion(4000000, 5000000, True) = True Then
            If chkConfirmHotKey.Value = 1 Then .dwFlags = .dwFlags Or TKF_CONFIRMHOTKEY
        End If
        If chkHotKeyActive.Value = 1 Then .dwFlags = .dwFlags Or TKF_HOTKEYACTIVE
        If chkHotKeySound.Value = 1 Then .dwFlags = .dwFlags Or TKF_HOTKEYSOUND
        If chkToggleKeysOn.Value = 1 Then .dwFlags = .dwFlags Or TKF_TOGGLEKEYSON
    End With
    
    If SystemParametersInfo(SPI_SETTOGGLEKEYS, TOGGLEKEYS.cbSize, TOGGLEKEYS, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim TOGGLEKEYS As TOGGLEKEYS
    TOGGLEKEYS.cbSize = Len(TOGGLEKEYS)
    
    If SystemParametersInfo(SPI_GETTOGGLEKEYS, TOGGLEKEYS.cbSize, TOGGLEKEYS, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    
    If TOGGLEKEYS.dwFlags And TKF_AVAILABLE Then
        With TOGGLEKEYS
            If .dwFlags And TKF_AVAILABLE Then chkAvailable.Value = 1
            If .dwFlags And TKF_HOTKEYACTIVE Then chkHotKeyActive.Value = 1
            If .dwFlags And TKF_HOTKEYSOUND Then chkHotKeySound.Value = 1
            If .dwFlags And TKF_TOGGLEKEYSON Then chkToggleKeysOn.Value = 1
            
            If WinVersion(4000000, 5000000, True) = True Then
                If .dwFlags And TKF_CONFIRMHOTKEY Then chkConfirmHotKey.Value = 1
            Else
                chkConfirmHotKey.Enabled = False
            End If
        End With
    Else
        lblHotKeyActive.Enabled = False
        chkHotKeyActive.Enabled = False
        lblHotKeySound.Enabled = False
        chkHotKeySound.Enabled = False
        lblToggleKeysOn.Enabled = False
        chkToggleKeysOn.Enabled = False
        lblConfirmHotKey.Enabled = False
        chkConfirmHotKey.Enabled = False
        cmdApply.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
