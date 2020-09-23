VERSION 5.00
Begin VB.Form frmNicnameWhois 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nicname/Whois"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmNicnameWhois.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   4575
   End
   Begin VB.CommandButton cmdSendData 
      Caption         =   "Send Data"
      Height          =   350
      Left            =   5280
      TabIndex        =   9
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtReturned 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1440
      Width           =   6135
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4320
      TabIndex        =   8
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblHostIP 
      Caption         =   "Host / IP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblReturned 
      Caption         =   "Returned"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmNicnameWhois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmNicnameWhois"


Private Sub cmdSendData_Click()
On Error GoTo VB_Error
    
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    
    cmdSendData.Enabled = False
    cmdStop.Enabled = True
    txtReturned.Text = vbNullString
    
    
    Call WSv4_Start(txtHostIP.Text, txtPort.Text, frmNicnameWhois.hwnd, 1, wsNicnameWhois_Socket, wsNicnameWhois_sockaddr)
    If send(wsNicnameWhois_Socket, ByVal txtSend.Text & vbCrLf, Len(txtSend.Text & vbCrLf), 0&) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdSendData_Click", "send")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdSendData_Click")
Resume Next
End Sub

Private Sub cmdStop_Click()
On Error GoTo VB_Error

    If shutdown(wsNicnameWhois_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdStop_Click", "shutdown")
    
    cmdStop.Enabled = False
    cmdSendData.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdStop_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Forms_Loaded.bNicnameWhois = True
    
    
    txtHostIP.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "HostIP")
    txtPort.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Port")
    txtSend.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Send")
    
    wsNicnameWhois_OldProc = SetWindowLong(frmNicnameWhois.hwnd, GWL_WNDPROC, AddressOf wsNicnameWhois_Proc): If wsNicnameWhois_OldProc = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SetWindowLong")
    
    
    If bWinsock = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblSend.Enabled = False
        txtSend.Enabled = False
        lblReturned.Enabled = False
        txtReturned.Enabled = False
        cmdStop.Enabled = False
        cmdSendData.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Forms_Loaded.bNicnameWhois = False
    
    
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "HostIP", txtHostIP.Text, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Port", txtPort.Text, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Send", txtSend.Text, REG_SZ)
    
    If wsNicnameWhois_Socket <> 0 Then
        If shutdown(wsNicnameWhois_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "shutdown")
        Call Socket_Close(wsNicnameWhois_Socket)
        
        Dim sockaddr_in As sockaddr_in
        wsNicnameWhois_sockaddr = sockaddr_in
    End If
    
    If SetWindowLong(frmNicnameWhois.hwnd, GWL_WNDPROC, wsNicnameWhois_OldProc) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "SetWindowLong")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
