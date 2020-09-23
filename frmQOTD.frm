VERSION 5.00
Begin VB.Form frmQOTD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quote Of The Day"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmQOTD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   350
      Left            =   2880
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1920
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtReturned 
      Height          =   1365
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
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
   Begin VB.Label lblMethod 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
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
   Begin VB.Label lblHostIP 
      Caption         =   "Host / IP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmQOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmQOTD"


Private Sub cmdGetData_Click()
On Error GoTo VB_Error
    
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    cmdGetData.Enabled = False
    cmdStop.Enabled = True
    txtReturned.Text = vbNullString
    
    
    Select Case cboMethod.ListIndex
        Case 0 'UDP
            Call WSv4_Start(txtHostIP.Text, txtPort.Text, frmQOTD.hwnd, 0, wsQOTD_Socket, wsQOTD_sockaddr)
            If sendto(wsQOTD_Socket, ByVal 0&, 0&, 0&, wsQOTD_sockaddr, Len(wsQOTD_sockaddr)) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdGetData_Click", "sendto")
        Case 1 'TCP
            Call WSv4_Start(txtHostIP.Text, txtPort.Text, frmQOTD.hwnd, 1, wsQOTD_Socket, wsQOTD_sockaddr)
    End Select
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdGetData_Click")
Resume Next
End Sub

Private Sub cmdStop_Click()
On Error GoTo VB_Error

    If shutdown(wsQOTD_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdStop_Click", "shutdown")
    
    cmdStop.Enabled = False
    cmdGetData.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdStop_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    Forms_Loaded.bQOTD = True
    
    
    With cboMethod
        .AddItem "UDP"
        .AddItem "TCP"
    End With
    
    txtHostIP.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\QOTD", "HostIP")
    cboMethod.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Method"), 0, 1)
    txtPort.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Port")
    
    wsQOTD_OldProc = SetWindowLong(frmQOTD.hwnd, GWL_WNDPROC, AddressOf wsQOTD_Proc): If wsQOTD_OldProc = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SetWindowLong")
    
    
    If bWinsock = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblMethod.Enabled = False
        cboMethod.Enabled = False
        lblReturned.Enabled = False
        txtReturned.Enabled = False
        cmdStop.Enabled = False
        cmdGetData.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Forms_Loaded.bQOTD = False
    
    
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\QOTD", "HostIP", txtHostIP.Text, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Method", cboMethod.ListIndex, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Port", txtPort.Text, REG_DWORD)
    
    If wsQOTD_Socket <> 0 Then
        If shutdown(wsQOTD_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "shutdown")
        Call Socket_Close(wsQOTD_Socket)
        
        Dim sockaddr_in As sockaddr_in
        wsQOTD_sockaddr = sockaddr_in
    End If
    
    If SetWindowLong(frmQOTD.hwnd, GWL_WNDPROC, wsQOTD_OldProc) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "SetWindowLong")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
