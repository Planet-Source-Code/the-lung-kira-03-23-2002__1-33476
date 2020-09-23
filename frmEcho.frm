VERSION 5.00
Begin VB.Form frmEcho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Echo"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmEcho.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "7"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox chkReturnOK 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtDataSize 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdSendData 
      Caption         =   "Send Data"
      Height          =   350
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1920
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblReturnOK 
      Caption         =   "Return OK"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblDataSize 
      Caption         =   "Data Size"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
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
   Begin VB.Label lblMethod 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmEcho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmEcho"


Private Sub cmdSendData_Click()
On Error GoTo VB_Error
    
    txtDataSize.Text = MinMax(Val(txtDataSize.Text), 1, 65536)
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    
    cmdSendData.Enabled = False
    cmdStop.Enabled = True
    chkReturnOK.value = 0
    
    
    wsEcho_Data = vbNullString
    
    
    wsEcho_Data = String$(txtDataSize.Text, 0)
    Select Case cboMethod.ListIndex
        Case 0 'UDP
            Call WSv4_Start(txtHostIP.Text, txtPort.Text, frmEcho.hwnd, 0, wsEcho_Socket, wsEcho_sockaddr)
            If sendto(wsEcho_Socket, ByVal wsEcho_Data, Len(wsEcho_Data), 0&, wsEcho_sockaddr, Len(wsEcho_sockaddr)) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdSendData_Click", "sendto")
        Case 1 'TCP
            Call WSv4_Start(txtHostIP.Text, txtPort.Text, frmEcho.hwnd, 1, wsEcho_Socket, wsEcho_sockaddr)
            If send(wsEcho_Socket, ByVal wsEcho_Data, Len(wsEcho_Data), 0&) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdSendData_Click", "send")
    End Select
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdSendData_Click")
Resume Next
End Sub

Private Sub cmdStop_Click()
On Error GoTo VB_Error

    If shutdown(wsEcho_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdStop_Click", "shutdown")
    
    chkReturnOK.value = 0
    cmdStop.Enabled = False
    cmdSendData.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdStop_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Forms_Loaded.bEcho = True
    
    
    With cboMethod
        .AddItem "UDP"
        .AddItem "TCP"
    End With
    
    
    txtDataSize.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Echo", "DataSize")
    txtHostIP.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Echo", "HostIP")
    cboMethod.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Echo", "Method"), 0, 1)
    txtPort.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Echo", "Port")
    
    wsEcho_OldProc = SetWindowLong(frmEcho.hwnd, GWL_WNDPROC, AddressOf wsEcho_Proc): If wsEcho_OldProc = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SetWindowLong")
    
    
    If bWinsock = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblMethod.Enabled = False
        cboMethod.Enabled = False
        lblDataSize.Enabled = False
        txtDataSize.Enabled = False
        lblReturnOK.Enabled = False
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
    
    Forms_Loaded.bEcho = False
    
    
    txtDataSize.Text = MinMax(Val(txtDataSize.Text), 1, 65536)
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "DataSize", txtDataSize.Text, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "HostIP", txtHostIP.Text, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "Method", cboMethod.ListIndex, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "Port", txtPort.Text, REG_DWORD)
    
    If wsEcho_Socket <> 0 Then
        If shutdown(wsEcho_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "shutdown")
        Call Socket_Close(wsEcho_Socket)
        
        wsEcho_Data = vbNullString
        
        Dim sockaddr_in As sockaddr_in
        wsEcho_sockaddr = sockaddr_in
    End If
    
    If SetWindowLong(frmEcho.hwnd, GWL_WNDPROC, wsEcho_OldProc) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "SetWindowLong")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
