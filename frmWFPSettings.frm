VERSION 5.00
Begin VB.Form frmWFPSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows File Protection - Settings"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmWFPSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQuota 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   2175
   End
   Begin VB.ComboBox cboWFP 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3000
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cboScan 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.CheckBox chkShowProgress 
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblQuota 
      Caption         =   "Quota"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblWFP 
      Caption         =   "WFP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblScan 
      Caption         =   "Scan"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblShowProgress 
      Caption         =   "Show Progress"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmWFPSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmWFPSettings"


Private Sub cboScan_Click()
On Error GoTo VB_Error
    
    If lblScan.Enabled = False Then lblScan.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboScan_Click")
Resume Next
End Sub

Private Sub cboWFP_Click()
On Error GoTo VB_Error
    
    If lblWFP.Enabled = False Then lblWFP.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboWFP_Click")
Resume Next
End Sub

Private Sub chkShowProgress_Click()
On Error GoTo VB_Error

    If lblShowProgress.Enabled = False Then lblShowProgress.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkShowProgress_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtQuota.Text = MinMax(Val(txtQuota.Text), 0, 4294967295#)
    
    
    If lblWFP.Enabled = True Then
        If cboWFP.ListIndex > -1 Then
            Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCDisable", cboWFP.ListIndex, REG_DWORD)
        End If
    End If
    If lblQuota.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCQuota", uint32_int32(txtQuota.Text), REG_DWORD)
    If lblScan.Enabled = True Then
        If cboScan.ListIndex > -1 Then
            Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCScan", cboScan.ListIndex, REG_DWORD)
        End If
    End If
    If lblShowProgress.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCShowProgress", chkShowProgress.value, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If WinVersion(-1, 5000000, True) = True Then
        With cboWFP
            .AddItem "Normal"
            .AddItem "Ask"
            .AddItem "Once"
            .AddItem "Setup"
            .AddItem "No PopUps"
        End With
        With cboScan
            .AddItem "Normal"
            .AddItem "Always"
            .AddItem "Once"
            .AddItem "Immediate"
        End With

        
        Dim bFail As Byte
        Dim lReturn As Long
        
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCDisable", bFail)
        If bFail <> 0 Then
            lblWFP.Enabled = False
        Else
            Select Case lReturn
                Case 0 To 4: cboWFP.ListIndex = lReturn
                Case Else: cboWFP.ListIndex = -1
            End Select
        End If
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCQuota", bFail)
        If bFail <> 0 Then
            lblQuota.Enabled = False
        Else
            txtQuota.Text = int32_uint32(lReturn)
        End If
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCScan", bFail)
        If bFail <> 0 Then
            lblScan.Enabled = False
        Else
            Select Case lReturn
                Case 0 To 3: cboScan.ListIndex = lReturn
                Case Else: cboScan.ListIndex = -1
            End Select
        End If
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCShowProgress", bFail)
        If bFail <> 0 Then
            lblShowProgress.Enabled = False
        Else
            chkShowProgress.value = MinMax(lReturn, 0, 1)
        End If
    Else
        lblWFP.Enabled = False
        cboWFP.Enabled = False
        lblQuota.Enabled = False
        txtQuota.Enabled = False
        lblScan.Enabled = False
        cboScan.Enabled = False
        lblShowProgress.Enabled = False
        chkShowProgress.Enabled = False
        cmdApply.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtQuota_Change")
Resume Next
End Sub

Private Sub txtQuota_Change()
On Error GoTo VB_Error

    If lblQuota.Enabled = False Then lblQuota.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtQuota_Change")
Resume Next
End Sub
