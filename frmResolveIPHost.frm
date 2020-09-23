VERSION 5.00
Begin VB.Form frmResolveIPHost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resolve IP Host"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmResolveIPHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetHost 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get Host"
      Height          =   350
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmdGetIP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get IP"
      Height          =   350
      Left            =   3960
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cboIP 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblHost 
      Caption         =   "Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblIP 
      Caption         =   "IP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmResolveIPHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmResolveIPHost"


Private Sub cmdGetHost_Click()
On Error GoTo VB_Error

    cmdGetHost.Enabled = False
    cmdGetIP.Enabled = False
    
    txtHost.Text = Host_IP(cboIP.Text)
    
    cmdGetIP.Enabled = True
    cmdGetHost.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdGetHost_Click")
Resume Next
End Sub

Private Sub cmdGetIP_Click()
On Error GoTo VB_Error

    cmdGetIP.Enabled = False
    cmdGetHost.Enabled = False
    
    With cboIP
        .Clear
        
        Dim asIP() As String
        Dim lCount As Long
        lCount = IP_Host(txtHost.Text, asIP())
        
        Dim lIncrement As Long
        For lIncrement = 0 To lCount
            If asIP(lIncrement) <> vbNullString Then
                .AddItem asIP(lIncrement)
            End If
        Next lIncrement
        
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    
    cmdGetHost.Enabled = True
    cmdGetIP.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdGetIP_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    txtHost.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "Host")
    
    Dim sReturn As String
    sReturn = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "IP")
    If sReturn <> vbNullString Then
        cboIP.AddItem sReturn
        cboIP.ListIndex = 0
    End If
    
    
    If bWinsock = False Then
        lblIP.Enabled = False
        cboIP.Enabled = False
        lblHost.Enabled = False
        txtHost.Enabled = False
        cmdGetIP.Enabled = False
        cmdGetHost.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "Host", txtHost.Text, REG_SZ)
    
    If cboIP.ListIndex > 0 Then
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "IP", cboIP.List(cboIP.ListIndex), REG_SZ)
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
