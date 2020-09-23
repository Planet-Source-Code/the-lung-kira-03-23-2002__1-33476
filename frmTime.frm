VERSION 5.00
Begin VB.Form frmTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox chkDaylightSavings 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton cmdSetTime 
      Caption         =   "Set Time"
      Height          =   350
      Left            =   2880
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtReturnedLocal 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   350
      Left            =   1800
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtReturnedGMT 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtUnFormatted 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   840
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblDaylightSavings 
      Caption         =   "Daylight Savings"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblReturnedLocal 
      Caption         =   "Returned Local"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
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
   Begin VB.Label lblReturnedGMT 
      Caption         =   "Returned GMT"
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
   Begin VB.Label lblUnFormatted 
      Caption         =   "UnFormatted"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmTime"


Private Sub cmdGetData_Click()
On Error GoTo VB_Error
    
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    
    cmdSetTime.Enabled = False
    cmdGetData.Enabled = False
    cmdStop.Enabled = True
    txtReturnedGMT.Text = vbNullString
    txtReturnedLocal.Text = vbNullString
    txtUnFormatted.Text = vbNullString
    
    
    Select Case cboMethod.ListIndex
        Case 0 'UDP
            Call WSv4_Start(txtHostIP.Text, txtPort.Text, frmTime.hwnd, 1, wsTime_Socket, wsTime_sockaddr)
            If sendto(wsTime_Socket, ByVal 0&, 0&, 0&, wsTime_sockaddr, Len(wsTime_sockaddr)) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdGetData_Click", "sendto")
        Case 1 'TCP
            Call WSv4_Start(txtHostIP.Text, txtPort.Text, frmTime.hwnd, 1, wsTime_Socket, wsTime_sockaddr)
    End Select
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdGetData_Click")
Resume Next
End Sub

Private Sub cmdSetTime_Click()
On Error GoTo VB_Error

    wsTime_SetTime = True
    Call cmdGetData_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdSetTime_Click")
Resume Next
End Sub

Private Sub cmdStop_Click()
On Error GoTo VB_Error

    If shutdown(wsTime_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdStop_Click", "shutdown")
    
    wsTime_SetTime = False
    
    cmdStop.Enabled = False
    cmdGetData.Enabled = True
    cmdSetTime.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdStop_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    Forms_Loaded.bTime = True
    
    
    With cboMethod
        .AddItem "UDP"
        .AddItem "TCP"
    End With
    
    chkDaylightSavings.value = IIf(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Time", "DaylightSavings"), 1, 0)
    txtHostIP.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Time", "HostIP")
    cboMethod.ListIndex = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Time", "Method")
    txtPort.Text = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Time", "Port")
    
    
    wsTime_OldProc = SetWindowLong(frmTime.hwnd, GWL_WNDPROC, AddressOf wsTime_Proc): If wsTime_OldProc = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SetWindowLong")
    
    
    If bWinsock = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblMethod.Enabled = False
        cboMethod.Enabled = False
        lblReturnedGMT.Enabled = False
        txtReturnedGMT.Enabled = False
        lblReturnedLocal.Enabled = False
        txtReturnedLocal.Enabled = False
        lblUnFormatted.Enabled = False
        txtUnFormatted.Enabled = False
        lblDaylightSavings.Enabled = False
        chkDaylightSavings.Enabled = False
        cmdStop.Enabled = False
        cmdGetData.Enabled = False
        cmdSetTime.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Forms_Loaded.bTime = False
    
    
    txtPort.Text = MinMax(Val(txtPort.Text), 0, 65535)
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "DaylightSavings", chkDaylightSavings.value, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "HostIP", txtHostIP.Text, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "Method", cboMethod.ListIndex, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "Port", txtPort.Text, REG_DWORD)
    
    If wsTime_Socket <> 0 Then
        If shutdown(wsTime_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "shutdown")
        Call Socket_Close(wsTime_Socket)
        
        wsTime_SetTime = False
        
        Dim sockaddr_in As sockaddr_in
        wsTime_sockaddr = sockaddr_in
    End If
    
    If SetWindowLong(frmTime.hwnd, GWL_WNDPROC, wsTime_OldProc) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Unload", "SetWindowLong")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
