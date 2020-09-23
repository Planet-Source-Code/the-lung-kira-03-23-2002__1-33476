VERSION 5.00
Begin VB.Form frmOwner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Owner"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmOwner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOwner 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtOrginization 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblOwner 
      Caption         =   "Owner"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblOrginization 
      Caption         =   "Orginization"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmOwner"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    If WinVersion(0, -1, True) = True Then
        If lblOwner.Enabled = True Then
            Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner", txtOwner.Text, REG_SZ)
        End If
        If lblOrginization.Enabled = True Then
            Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", txtOrginization.Text, REG_SZ)
        End If
    Else
        If lblOwner.Enabled = True Then
            Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", txtOwner.Text, REG_SZ)
        End If
        If lblOrginization.Enabled = True Then
            Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", txtOrginization.Text, REG_SZ)
        End If
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    Dim bFail As Byte
    
    If WinVersion(0, -1, True) = True Then
        txtOwner.Text = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner", bFail)
        If bFail <> 0 Then lblOwner.Enabled = False
        
        txtOrginization.Text = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", bFail)
        If bFail <> 0 Then lblOrginization.Enabled = False
    Else
        txtOwner.Text = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", bFail)
        If bFail <> 0 Then lblOwner.Enabled = False
        
        txtOrginization.Text = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", bFail)
        If bFail <> 0 Then lblOrginization.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub txtOrginization_Change()
On Error GoTo VB_Error
    
    If lblOrginization.Enabled = False Then lblOrginization.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtOrginization_Change")
Resume Next
End Sub

Private Sub txtOwner_Change()
On Error GoTo VB_Error
    
    If lblOwner.Enabled = False Then lblOwner.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtOwner_Change")
Resume Next
End Sub
