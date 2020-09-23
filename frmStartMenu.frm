VERSION 5.00
Begin VB.Form frmStartMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Start Menu"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmStartMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkStartBanner 
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkHelp 
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox chkNetworkConnections 
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox chkRun 
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chkFind 
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkRecentDocsMenu 
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox chkRecentDocsHistory 
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chkLogoff 
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2400
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox chkFavoritesMenu 
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblStartBanner 
      Caption         =   "Start Banner"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label lblHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1965
   End
   Begin VB.Label lblNetworkConnections 
      Caption         =   "Network Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1965
   End
   Begin VB.Label lblRun 
      Caption         =   "Run"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   1965
   End
   Begin VB.Label lblFind 
      Caption         =   "Find"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1965
   End
   Begin VB.Label lblRecentDocsMenu 
      Caption         =   "Recent Docs Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1965
   End
   Begin VB.Label lblRecentDocsHistory 
      Caption         =   "Recent Docs History"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1965
   End
   Begin VB.Label lblLogoff 
      Caption         =   "Logoff"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1965
   End
   Begin VB.Label lblFavoritesMenu 
      Caption         =   "Favorites Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1965
   End
End
Attribute VB_Name = "frmStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmStartMenu"


Private Sub chkFavoritesMenu_Click()
On Error GoTo VB_Error

    If lblFavoritesMenu.Enabled = False Then lblFavoritesMenu.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkFavoritesMenu_Click")
Resume Next
End Sub

Private Sub chkFind_Click()
On Error GoTo VB_Error

    If lblFind.Enabled = False Then lblFind.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkFind_Click")
Resume Next
End Sub

Private Sub chkHelp_Click()
On Error GoTo VB_Error

    If lblHelp.Enabled = False Then lblHelp.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkHelp_Click")
Resume Next
End Sub

Private Sub chkLogoff_Click()
On Error GoTo VB_Error

    If lblLogoff.Enabled = False Then lblLogoff.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkLogoff_Click")
Resume Next
End Sub

Private Sub chkNetworkConnections_Click()
On Error GoTo VB_Error

    If lblNetworkConnections.Enabled = False Then lblNetworkConnections.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkNetworkConnections_Click")
Resume Next
End Sub

Private Sub chkRecentDocsHistory_Click()
On Error GoTo VB_Error

    If lblRecentDocsHistory.Enabled = False Then lblRecentDocsHistory.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkRecentDocsHistory_Click")
Resume Next
End Sub

Private Sub chkRecentDocsMenu_Click()
On Error GoTo VB_Error

    If lblRecentDocsMenu.Enabled = False Then lblRecentDocsMenu.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkRecentDocsMenu_Click")
Resume Next
End Sub

Private Sub chkRun_Click()
On Error GoTo VB_Error

    If lblRun.Enabled = False Then lblRun.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkRun_Click")
Resume Next
End Sub

Private Sub chkStartBanner_Click()
On Error GoTo VB_Error

    If lblStartBanner.Enabled = False Then lblStartBanner.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkStartBanner_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error

    If lblStartBanner.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoStartBanner", IIf(chkStartBanner.value, 0, 1), REG_BINARY)
    End If
    
    If lblFavoritesMenu.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", IIf(chkFavoritesMenu.value, 0, 1), REG_BINARY)
    End If
    
    If lblFind.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", IIf(chkFind.value, 0, 1), REG_BINARY)
    End If
    
    If lblHelp.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", IIf(chkHelp.value, 0, 1), REG_BINARY)
    End If
    
    If lblLogoff.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", IIf(chkLogoff.value, 0, 1), REG_BINARY)
    End If
    
    If lblNetworkConnections.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetworkConnections", IIf(chkNetworkConnections.value, 0, 1), REG_BINARY)
    End If
    
    If lblRecentDocsHistory.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory", IIf(chkRecentDocsHistory.value, 0, 1), REG_BINARY)
    End If
    
    If lblRecentDocsMenu.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", IIf(chkRecentDocsMenu.value, 0, 1), REG_BINARY)
    End If
    
    If lblRun.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", IIf(chkRun.value, 0, 1), REG_BINARY)
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim bFail As Byte
    
    If Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoStartBanner", bFail) = 0 Then chkStartBanner.value = 1
    If bFail <> 0 Then lblStartBanner.Enabled = False
    
    If Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", bFail) = 0 Then chkFavoritesMenu.value = 1
    If bFail <> 0 Then lblFavoritesMenu.Enabled = False
    
    If Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", bFail) = 0 Then chkFind.value = 1
    If bFail <> 0 Then lblFind.Enabled = False
    
    If Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", bFail) = 0 Then chkHelp.value = 1
    If bFail <> 0 Then lblHelp.Enabled = False
    
    If Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", bFail) = 0 Then chkLogoff.value = 1
    If bFail <> 0 Then lblLogoff.Enabled = False
    
    If Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetworkConnections", bFail) = 0 Then chkNetworkConnections.value = 1
    If bFail <> 0 Then lblNetworkConnections.Enabled = False
    
    If Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory", bFail) = 0 Then chkRecentDocsHistory.value = 1
    If bFail <> 0 Then lblRecentDocsHistory.Enabled = False
    
    If Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", bFail) = 0 Then chkRecentDocsMenu.value = 1
    If bFail <> 0 Then lblRecentDocsMenu.Enabled = False
    
    If Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", bFail) = 0 Then chkRun.value = 1
    If bFail <> 0 Then lblRun.Enabled = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
