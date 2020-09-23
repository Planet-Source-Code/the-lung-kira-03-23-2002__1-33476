VERSION 5.00
Begin VB.Form frmMIB2IFTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MIB-II Interface Table"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmMIB2IFTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbps 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "bps"
      Top             =   3600
      Width           =   255
   End
   Begin VB.TextBox txtSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtPhysicalAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtOutputQueueLength 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtOperationalStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtMTU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtLastStatusChange 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtInterfaceType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtInterfaceIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtErroneousOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtDiscardedOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtNonUnicastOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtUnicastOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtOctetsOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtUnknownProtocolIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtErroneousIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtDiscardedIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtNonUnicastIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtUnicastIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtOctetsIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtAdminStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   6840
      TabIndex        =   50
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   350
      Left            =   6120
      TabIndex        =   51
      Top             =   3840
      Width           =   975
   End
   Begin VB.ListBox lstMIB2IF_Table 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblOutputQueueLength 
      Caption         =   "Output Queue Length"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblPacketsOut 
      Caption         =   "Packets Out"
      Height          =   255
      Left            =   3720
      TabIndex        =   38
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblErroneousOut 
      Caption         =   "Erroneous"
      Height          =   255
      Left            =   3720
      TabIndex        =   47
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblDiscardedOut 
      Caption         =   "Discarded"
      Height          =   255
      Left            =   3720
      TabIndex        =   45
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblNonUnicastOut 
      Caption         =   "Non Unicast"
      Height          =   255
      Left            =   3720
      TabIndex        =   43
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblUnicastOut 
      Caption         =   "Unicast"
      Height          =   255
      Left            =   3720
      TabIndex        =   41
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblOctetsOut 
      Caption         =   "Octets"
      Height          =   255
      Left            =   3720
      TabIndex        =   39
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblPacketsIn 
      Caption         =   "Packets In"
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblUnknownProtocolIn 
      Caption         =   "Unknown Protocol"
      Height          =   255
      Left            =   3720
      TabIndex        =   36
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblErroneousIn 
      Caption         =   "Erroneous"
      Height          =   255
      Left            =   3720
      TabIndex        =   34
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblDiscardedIn 
      Caption         =   "Discarded"
      Height          =   255
      Left            =   3720
      TabIndex        =   32
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblNonUnicastIn 
      Caption         =   "Non Unicast"
      Height          =   255
      Left            =   3720
      TabIndex        =   30
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblUnicastIn 
      Caption         =   "Unicast"
      Height          =   255
      Left            =   3720
      TabIndex        =   28
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblOctetsIn 
      Caption         =   "Octets"
      Height          =   255
      Left            =   3720
      TabIndex        =   26
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblLastStatusChange 
      Caption         =   "Last Status Change"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblOperationalStatus 
      Caption         =   "Operational Status"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblPhysicalAddress 
      Caption         =   "Physical Address"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblAdminStatus 
      Caption         =   "Admin Status"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   3720
      TabIndex        =   49
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblMTU 
      Caption         =   "MTU"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblInterfaceType 
      Caption         =   "Interface Type"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblInterfaceIndex 
      Caption         =   "Interface Index"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblEntry 
      Caption         =   "Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMIB2IFTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim MIB_IFTABLE As MIB_IFTABLE
Const sLocation As String = "frmMIB2IFTable"


Private Sub cmdRefresh_Click()
On Error GoTo VB_Error

    Dim lSize As Long
    
    lSize = Len(MIB_IFTABLE)
    lErrors = GetIfTable(MIB_IFTABLE, lSize, chkSorted.value): If lErrors <> NO_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "GetIfTable")
    
    
    With lstMIB2IF_Table
        .Clear
        
        Dim lIncrement As Long
        For lIncrement = 0 To MIB_IFTABLE.dwNumEntries - 1
            .AddItem (lIncrement + 1)
        Next lIncrement
    End With
    
    txtAdminStatus.Text = vbNullString
    txtDescription.Text = vbNullString
    txtDiscardedIn.Text = vbNullString
    txtDiscardedOut.Text = vbNullString
    txtErroneousIn.Text = vbNullString
    txtErroneousOut.Text = vbNullString
    txtName.Text = vbNullString
    txtInterfaceIndex.Text = vbNullString
    txtInterfaceType.Text = vbNullString
    txtLastStatusChange.Text = vbNullString
    txtMTU.Text = vbNullString
    txtNonUnicastIn.Text = vbNullString
    txtNonUnicastOut.Text = vbNullString
    txtOctetsIn.Text = vbNullString
    txtOctetsOut.Text = vbNullString
    txtOutputQueueLength.Text = vbNullString
    txtOperationalStatus.Text = vbNullString
    txtPhysicalAddress.Text = vbNullString
    txtSpeed.Text = vbNullString
    txtUnicastIn.Text = vbNullString
    txtUnicastOut.Text = vbNullString
    txtUnknownProtocolIn.Text = vbNullString
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdRefresh_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("iphlpapi.dll", "GetIfTable") = True Then
        Call cmdRefresh_Click
    Else
        lblEntry.Enabled = False
        lstMIB2IF_Table.Enabled = False
        lblAdminStatus.Enabled = False
        lblDescription.Enabled = False
        lblInterfaceIndex.Enabled = False
        lblInterfaceType.Enabled = False
        lblLastStatusChange.Enabled = False
        lblMTU.Enabled = False
        lblName.Enabled = False
        lblOperationalStatus.Enabled = False
        lblOutputQueueLength.Enabled = False
        lblPhysicalAddress.Enabled = False
        lblSpeed.Enabled = False
        lblPacketsIn.Enabled = False
        lblOctetsIn.Enabled = False
        lblUnicastIn.Enabled = False
        lblNonUnicastIn.Enabled = False
        lblDiscardedIn.Enabled = False
        lblErroneousIn.Enabled = False
        lblUnknownProtocolIn.Enabled = False
        lblPacketsOut.Enabled = False
        lblOctetsOut.Enabled = False
        lblUnicastOut.Enabled = False
        lblNonUnicastOut.Enabled = False
        lblDiscardedOut.Enabled = False
        lblErroneousOut.Enabled = False
        lblSorted.Enabled = False
        cmdRefresh.Enabled = False
    End If
    
    chkSorted.value = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MIB2IFTable", "Sorted")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MIB2IFTable", "Sorted", chkSorted.value, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub lstMIB2IF_Table_Click()
On Error GoTo VB_Error

    With MIB_IFTABLE.table(lstMIB2IF_Table.ListIndex)
        txtName.Text = Unicode_Ascii(.wszName, 0&)
        txtInterfaceIndex.Text = int32_uint32(.dwIndex)
        
        Select Case .dwType
            Case MIB_IF_TYPE_OTHER: txtInterfaceType.Text = "Other"
            Case MIB_IF_TYPE_ETHERNET: txtInterfaceType.Text = "Ethernet"
            Case MIB_IF_TYPE_TOKENRING: txtInterfaceType.Text = "Tokenring"
            Case MIB_IF_TYPE_FDDI: txtInterfaceType.Text = "FDDI"
            Case MIB_IF_TYPE_PPP: txtInterfaceType.Text = "PPP"
            Case MIB_IF_TYPE_LOOPBACK: txtInterfaceType.Text = "Loopback"
            Case MIB_IF_TYPE_SLIP: txtInterfaceType.Text = "Slip"
            Case Else: txtInterfaceType.Text = "Unknown " & int32_uint32(.dwType)
        End Select
        
        txtMTU.Text = FormatNumber(int32_uint32(.dwMtu), 0, , , True)
        txtSpeed.Text = FormatNumber(int32_uint32(.dwSpeed), 0, , , True)
        
        If Len(.bPhysAddr) >= .dwPhysAddrLen Then
            Dim strAddress As String
            Dim lIncrement As Long
            
            strAddress = Left$(.bPhysAddr, .dwPhysAddrLen)
            txtPhysicalAddress.Text = vbNullString
            
            For lIncrement = 1 To .dwPhysAddrLen
                txtPhysicalAddress.Text = txtPhysicalAddress.Text & ltoa_(Asc(Mid$(.bPhysAddr, lIncrement, 1)), 16)
            Next lIncrement
        End If
        
        Select Case .dwAdminStatus
            Case MIB_IF_ADMIN_STATUS_UP: txtAdminStatus.Text = "Up"
            Case MIB_IF_ADMIN_STATUS_DOWN: txtAdminStatus.Text = "Down"
            Case MIB_IF_ADMIN_STATUS_TESTING: txtAdminStatus.Text = "Testing"
            Case Else: txtAdminStatus.Text = "Unknown " & int32_uint32(.dwAdminStatus)
        End Select
        
        Select Case .dwOperStatus
            Case MIB_IF_OPER_STATUS_NON_OPERATIONAL: txtOperationalStatus.Text = "Non Operational"
            Case MIB_IF_OPER_STATUS_UNREACHABLE: txtOperationalStatus.Text = "Unreachable"
            Case MIB_IF_OPER_STATUS_DISCONNECTED: txtOperationalStatus.Text = "Disconnected"
            Case MIB_IF_OPER_STATUS_CONNECTING: txtOperationalStatus.Text = "Connecting"
            Case MIB_IF_OPER_STATUS_CONNECTED: txtOperationalStatus.Text = "Connected"
            Case MIB_IF_OPER_STATUS_OPERATIONAL: txtOperationalStatus.Text = "Operational"
            Case Else: txtOperationalStatus.Text = "Unknown " & int32_uint32(.dwOperStatus)
        End Select
        
        txtLastStatusChange.Text = FormatNumber(int32_uint32(.dwLastChange), 0, , , True)
        txtOctetsIn.Text = FormatNumber(int32_uint32(.dwInOctets), 0, , , True)
        txtUnicastIn.Text = FormatNumber(int32_uint32(.dwInUcastPkts), 0, , , True)
        txtNonUnicastIn.Text = FormatNumber(int32_uint32(.dwInNUcastPkts), 0, , , True)
        txtDiscardedIn.Text = FormatNumber(int32_uint32(.dwInDiscards), 0, , , True)
        txtErroneousIn.Text = FormatNumber(int32_uint32(.dwInErrors), 0, , , True)
        txtUnknownProtocolIn.Text = FormatNumber(int32_uint32(.dwInUnknownProtos), 0, , , True)
        txtOctetsOut.Text = FormatNumber(int32_uint32(.dwOutOctets), 0, , , True)
        txtUnicastOut.Text = FormatNumber(int32_uint32(.dwOutUcastPkts), 0, , , True)
        txtNonUnicastOut.Text = FormatNumber(int32_uint32(.dwOutNUcastPkts), 0, , , True)
        txtDiscardedOut.Text = FormatNumber(int32_uint32(.dwOutDiscards), 0, , , True)
        txtErroneousOut.Text = FormatNumber(int32_uint32(.dwOutErrors), 0, , , True)
        txtOutputQueueLength.Text = FormatNumber(int32_uint32(.dwOutQLen), 0, , , True)
        
        If Len(.bDescr) >= .dwDescrLen Then
            txtDescription.Text = Left$(.bDescr, .dwDescrLen)
        End If
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstMIB2IF_Table_Click")
Resume Next
End Sub
