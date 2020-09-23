VERSION 5.00
Begin VB.Form frmTCPTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP Table"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmTCPTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConnectionState 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtRemotePort 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtLocalPort 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtRemoteAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtLocalAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2400
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox lstTCP_Table 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apply"
      Height          =   350
      Left            =   1320
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblRemotePort 
      Caption         =   "Remote Port"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblRemoteAddress 
      Caption         =   "Remote Address"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblLocalPort 
      Caption         =   "Local Port"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblLocalAddress 
      Caption         =   "Local Address"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblConnectionState 
      Caption         =   "Connection State"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblEntry 
      Caption         =   "Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmTCPTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim MIB_TCPTABLE As MIB_TCPTABLE
Const sLocation As String = "frmTCPTable"


Private Sub cmdRefresh_Click()
On Error GoTo VB_Error

    Dim lSize As Long
    lSize = Len(MIB_TCPTABLE)
    
    lErrors = GetTcpTable(MIB_TCPTABLE, lSize, chkSorted.value): If lErrors <> NO_ERROR Then Call Error_API(lErrors, sLocation & "\cmdRefresh_Click", "GetTcpTable")
    
    
    With lstTCP_Table
        .Clear
        
        Dim lIncrement As Long
        For lIncrement = 0 To MIB_TCPTABLE.dwNumEntries - 1
            .AddItem (lIncrement + 1)
        Next lIncrement
    End With
    
    txtConnectionState.Text = vbNullString
    txtLocalAddress.Text = vbNullString
    txtLocalPort.Text = vbNullString
    txtRemoteAddress.Text = vbNullString
    txtRemotePort.Text = vbNullString
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdRefresh_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("iphlpapi.dll", "GetTcpTable") = True Then
        Call cmdRefresh_Click
    Else
        lblEntry.Enabled = False
        lstTCP_Table.Enabled = False
        lblConnectionState.Enabled = False
        lblLocalAddress.Enabled = False
        lblLocalPort.Enabled = False
        lblRemoteAddress.Enabled = False
        lblRemotePort.Enabled = False
        lblSorted.Enabled = False
        chkSorted.Enabled = False
        cmdApply.Enabled = False
        cmdRefresh.Enabled = False
    End If
    
    chkSorted.value = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\TCPTable", "Sorted")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\TCPTable", "Sorted", chkSorted.value, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub lstTCP_Table_Click()
On Error GoTo VB_Error
    
    With MIB_TCPTABLE.table(lstTCP_Table.ListIndex)
        Select Case .dwState
            Case MIB_TCP_STATE_CLOSED: txtConnectionState.Text = "Closed"
            Case MIB_TCP_STATE_LISTEN: txtConnectionState.Text = "Listen"
            Case MIB_TCP_STATE_SYN_SENT: txtConnectionState.Text = "Syn Sent"
            Case MIB_TCP_STATE_SYN_RCVD: txtConnectionState.Text = "Syn Received"
            Case MIB_TCP_STATE_ESTAB: txtConnectionState.Text = "Established"
            Case MIB_TCP_STATE_FIN_WAIT1: txtConnectionState.Text = "Finished Wait1"
            Case MIB_TCP_STATE_FIN_WAIT2: txtConnectionState.Text = "Finished Wait2"
            Case MIB_TCP_STATE_CLOSE_WAIT: txtConnectionState.Text = "Close Wait"
            Case MIB_TCP_STATE_CLOSING: txtConnectionState.Text = "Closing"
            Case MIB_TCP_STATE_LAST_ACK: txtConnectionState.Text = "Last Acknowledge"
            Case MIB_TCP_STATE_TIME_WAIT: txtConnectionState.Text = "Time Wait"
            Case MIB_TCP_STATE_DELETE_TCB: txtConnectionState.Text = "Delete TCB"
            Case Else: txtConnectionState.Text = "Unknown " & int32_uint32(.dwState)
        End Select
        
        txtLocalAddress.Text = IP_String(.dwLocalAddr)
        txtLocalPort.Text = htons(.dwLocalPort)
        txtRemoteAddress.Text = IP_String(.dwRemoteAddr)
        txtRemotePort.Text = htons(.dwRemotePort)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstTCP_Table_Click")
Resume Next
End Sub
