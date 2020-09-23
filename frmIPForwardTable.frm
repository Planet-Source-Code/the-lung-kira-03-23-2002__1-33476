VERSION 5.00
Begin VB.Form frmIPForwardTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Forward Table"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmIPForwardTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtForwardMetric5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtForwardMetric4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtForwardMetric3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtForwardMetric2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtRouteType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtProtocol 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtPolicy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtInterfaceIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtASNextHop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtNextHop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtForwardMetric1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "S"
      Top             =   2280
      Width           =   135
   End
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   3240
      TabIndex        =   32
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtDestination 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox lstIPForward_Table 
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
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2520
      TabIndex        =   33
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label lblForwardMetric5 
      Caption         =   "Forward Metric 5"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblForwardMetric4 
      Caption         =   "Forward Metric 4"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblForwardMetric3 
      Caption         =   "Forward Metric 3"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblForwardMetric2 
      Caption         =   "Forward Metric 2"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblForwardMetric1 
      Caption         =   "Forward Metric 1"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblASNextHop 
      Caption         =   "Next Hop System #"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblAge 
      Caption         =   "Age"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblProtocol 
      Caption         =   "Protocol"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblRouteType 
      Caption         =   "Route Type"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblInterfaceIndex 
      Caption         =   "Interface Index"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblNextHop 
      Caption         =   "Next Hop"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblPolicy 
      Caption         =   "Policy"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblMask 
      Caption         =   "Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblDestination 
      Caption         =   "Destination"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblEntry 
      Caption         =   "Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmIPForwardTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim MIB_IPFORWARDTABLE As MIB_IPFORWARDTABLE
Const sLocation As String = "frmIPForwardTable"


Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    Dim lSize As Long
    lSize = Len(MIB_IPFORWARDTABLE)
    
    lErrors = GetIpForwardTable(MIB_IPFORWARDTABLE, lSize, chkSorted.value): If lErrors <> NO_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "GetIpForwardTable")
    
    
    With lstIPForward_Table
        .Clear
        
        Dim lIncrement As Long
        For lIncrement = 0 To MIB_IPFORWARDTABLE.dwNumEntries - 1
            .AddItem (lIncrement + 1)
        Next lIncrement
    End With
    
    txtDestination.Text = vbNullChar
    txtMask.Text = vbNullChar
    txtPolicy.Text = vbNullChar
    txtNextHop.Text = vbNullChar
    txtInterfaceIndex.Text = vbNullChar
    txtRouteType.Text = vbNullChar
    txtProtocol.Text = vbNullChar
    txtAge.Text = vbNullChar
    txtASNextHop.Text = vbNullChar
    txtForwardMetric1.Text = vbNullChar
    txtForwardMetric2.Text = vbNullChar
    txtForwardMetric3.Text = vbNullChar
    txtForwardMetric4.Text = vbNullChar
    txtForwardMetric5.Text = vbNullChar
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdRefresh_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("iphlpapi.dll", "GetIpForwardTable") = True Then
        Call cmdRefresh_Click
    Else
        lblEntry.Enabled = False
        lstIPForward_Table.Enabled = False
        lblDestination.Enabled = False
        lblMask.Enabled = False
        lblPolicy.Enabled = False
        lblNextHop.Enabled = False
        lblInterfaceIndex.Enabled = False
        lblRouteType.Enabled = False
        lblProtocol.Enabled = False
        lblAge.Enabled = False
        lblASNextHop.Enabled = False
        lblForwardMetric1.Enabled = False
        lblForwardMetric2.Enabled = False
        lblForwardMetric3.Enabled = False
        lblForwardMetric4.Enabled = False
        lblForwardMetric5.Enabled = False
        lblSorted.Enabled = False
        chkSorted.Enabled = False
        cmdRefresh.Enabled = False
    End If

    chkSorted.value = Reg_Read(HKEY_CURRENT_USER, sRegKey & "\IPForwardTable", "Sorted")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\IPForwardTable", "Sorted", chkSorted.value, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub lstIPForward_Table_Click()
On Error GoTo VB_Error

    With MIB_IPFORWARDTABLE.table(lstIPForward_Table.ListIndex)
        txtDestination.Text = IP_String(.dwForwardDest)
        txtMask.Text = IP_String(.dwForwardMask)
        txtPolicy.Text = int32_uint32(.dwForwardPolicy)
        txtNextHop.Text = IP_String(.dwForwardNextHop)
        txtInterfaceIndex.Text = int32_uint32(.dwForwardIfIndex)
        
        Select Case .dwForwardType
            Case 4: txtRouteType.Text = "Not Final Destination"
            Case 3: txtRouteType.Text = "Final Destination"
            Case 2: txtRouteType.Text = "Invalid"
            Case 1: txtRouteType.Text = "Other"
            Case Else: txtRouteType.Text = "Unknown " & int32_uint32(.dwForwardType)
        End Select
        
        Select Case .dwForwardProto
            Case PROTO_IP_OTHER: txtProtocol.Text = "Other"
            Case PROTO_IP_LOCAL: txtProtocol.Text = "Local"
            Case PROTO_IP_NETMGMT: txtProtocol.Text = "NetMgmt"
            Case PROTO_IP_ICMP: txtProtocol.Text = "ICMP"
            Case PROTO_IP_EGP: txtProtocol.Text = "EGP"
            Case PROTO_IP_GGP: txtProtocol.Text = "GGP"
            Case PROTO_IP_HELLO: txtProtocol.Text = "HELLO"
            Case PROTO_IP_RIP: txtProtocol.Text = "RIP"
            Case PROTO_IP_IS_IS: txtProtocol.Text = "IS IS"
            Case PROTO_IP_ES_IS: txtProtocol.Text = "ES IS"
            Case PROTO_IP_CISCO: txtProtocol.Text = "Cisco"
            Case PROTO_IP_BBN: txtProtocol.Text = "BBN"
            Case PROTO_IP_OSPF: txtProtocol.Text = "OSPF"
            Case PROTO_IP_BGP: txtProtocol.Text = "BGP"
            Case PROTO_IP_BOOTP: txtProtocol.Text = "BootP"
            Case PROTO_IP_NT_AUTOSTATIC: txtProtocol.Text = "NT AutoStatic"
            Case PROTO_IP_NT_STATIC: txtProtocol.Text = "NT Static"
            Case PROTO_IP_NT_STATIC_NON_DOD: txtProtocol.Text = "NT Static Non DOD"
            Case Else: txtProtocol.Text = "Unknown " & int32_uint32(.dwForwardProto)
        End Select

        txtAge.Text = int32_uint32(.dwForwardAge)
        txtASNextHop.Text = int32_uint32(.dwForwardNextHopAS)
        
        txtForwardMetric1.Text = int32_uint32(.dwForwardMetric1)
        txtForwardMetric2.Text = int32_uint32(.dwForwardMetric2)
        txtForwardMetric3.Text = int32_uint32(.dwForwardMetric3)
        txtForwardMetric4.Text = int32_uint32(.dwForwardMetric4)
        txtForwardMetric5.Text = int32_uint32(.dwForwardMetric5)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstIPForward_Table_Click")
Resume Next
End Sub
