VERSION 5.00
Begin VB.Form frmDayTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DayTime"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmDayTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "13"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtReturned 
      Height          =   525
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1440
      Width           =   3735
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   350
      Left            =   2880
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
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
   Begin VB.Label lblMethod 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmDayTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmDayTime"


Private Sub cmdGetData_Click()
On Error GoTo VB_Error
    
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    
    cmdGetData.Enabled = False
    cmdStop.Enabled = True
    txtReturned.Text = vbNullString
    
    
    Select Case cboMethod.ListIndex
        Case 0 'UDP
            Call WSv4_Start(txtHostIP.Text, txtPort.Text, frmDayTime.hwnd, 0, wsDayTime_Socket, wsDayTime_sockaddr)
            If sendto(wsDayTime_Socket, ByVal 0&, 0&, 0&, wsDayTime_sockaddr, Len(wsDayTime_sockaddr)) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdGetData_Click", "sendto")
        Case 1 'TCP
            Call WSv4_Start(txtHostIP.Text, txtPort.Text, frmDayTime.hwnd, 1, wsDayTime_Socket, wsDayTime_sockaddr)
    End Select
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdGetData_Click")
Resume Next
End Sub

Private Sub cmdStop_Click()
On Error GoTo VB_Error

    If shutdown(wsDayTime_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdStop_Click", "shutdown")
    
    cmdStop.Enabled = False
    cmdGetData.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdStop_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    Forms_Loaded.bDayTime = True
    
    
    With cboMethod
        .AddItem "UDP"
        .AddItem "TCP"
    End With
    
    txtHostIP.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DayTime", "HostIP")
    cboMethod.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Method"), 0, 1)
    txtPort.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Port")
    
    wsDayTime_OldProc = SetWindowLong(frmDayTime.hwnd, GWL_WNDPROC, AddressOf wsDayTime_Proc): If wsDayTime_OldProc = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SetWindowLong")
    
    
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

    Forms_Loaded.bDayTime = False
    
    
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DayTime", "HostIP", txtHostIP.Text, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Method", cboMethod.ListIndex, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Port", txtPort.Text, REG_DWORD)
    
    If wsDayTime_Socket <> 0 Then
        If shutdown(wsDayTime_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "shutdown")
        Call Socket_Close(wsDayTime_Socket)
        
        Dim sockaddr_in As sockaddr_in
        wsDayTime_sockaddr = sockaddr_in
    End If
    
    If SetWindowLong(frmDayTime.hwnd, GWL_WNDPROC, wsDayTime_OldProc) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "SetWindowLong")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
