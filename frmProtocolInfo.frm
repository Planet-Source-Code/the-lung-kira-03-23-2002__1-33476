VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProtocolInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protocol Info"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmProtocolInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtSocketType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtSecurityScheme 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox txtReserved 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtProtocol 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtNetworkByteOrder 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtMinSocketAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtMaxSocketAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtMaxProtocolOffset 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtMaxMessageSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtServiceFlags4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txtServiceFlags3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtChainLength 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1680
      Width           =   1455
   End
   Begin VB.ListBox lstChainEntries 
      Height          =   645
      Left            =   3600
      TabIndex        =   26
      Top             =   2280
      Width           =   3255
   End
   Begin MSComctlLib.ListView lvwServiceFlags1 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtServiceFlags2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtAddressFamily 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CheckBox chkMatchesProtocol0 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox chkHidden 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox chkRecommendedEntry 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   5520
      Width           =   255
   End
   Begin VB.CheckBox chkMultipleEntries 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   5280
      Width           =   255
   End
   Begin VB.TextBox txtCatalogEntryID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtProviderID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1080
      Width           =   3255
   End
   Begin VB.ComboBox cboProtocols 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label lblChainEntries 
      Caption         =   "Chain Entries"
      Height          =   255
      Left            =   3600
      TabIndex        =   25
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblServiceFlags1 
      Caption         =   "Service Flags1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblServiceFlags2 
      Caption         =   "Service Flags 2"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblServiceFlags3 
      Caption         =   "Service Flags 3"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblServiceFlags4 
      Caption         =   "Service Flags 4"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblReserved 
      Caption         =   "Reserved"
      Height          =   255
      Left            =   3600
      TabIndex        =   41
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblMaxMessageSize 
      Caption         =   "Max Message Size"
      Height          =   255
      Left            =   3600
      TabIndex        =   29
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblSecurityScheme 
      Caption         =   "Security Scheme"
      Height          =   255
      Left            =   3600
      TabIndex        =   43
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblNetworkByteOrder 
      Caption         =   "Network Byte Order"
      Height          =   255
      Left            =   3600
      TabIndex        =   37
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblMaxProtocolOffset 
      Caption         =   "Max Protocol Offset"
      Height          =   255
      Left            =   3600
      TabIndex        =   31
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblProtocol 
      Caption         =   "Protocol"
      Height          =   255
      Left            =   3600
      TabIndex        =   39
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblSocketType 
      Caption         =   "Socket Type"
      Height          =   255
      Left            =   3600
      TabIndex        =   45
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblMaxSocketAddress 
      Caption         =   "Max Socket Address"
      Height          =   255
      Left            =   3600
      TabIndex        =   33
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblMinSocketAddress 
      Caption         =   "Min Socket Address"
      Height          =   255
      Left            =   3600
      TabIndex        =   35
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblAddressFamily 
      Caption         =   "Address Family"
      Height          =   255
      Left            =   3600
      TabIndex        =   27
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   255
      Left            =   3600
      TabIndex        =   47
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblMatchesProtocol0 
      Caption         =   "Matches Protocol 0"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblHidden 
      Caption         =   "Hidden"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblRecommendedEntry 
      Caption         =   "Recommended Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblMultipleEntries 
      Caption         =   "Multiple Entries"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblProviderFlags 
      Caption         =   "Provider Flags"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblChainLength 
      Caption         =   "Chain Length"
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblCatalogEntryID 
      Caption         =   "Catalog Entry ID"
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblProviderID 
      Caption         =   "Provider ID"
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblProtocols 
      Caption         =   "Protocols"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmProtocolInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim WSAPROTOCOL_INFO() As WSAPROTOCOL_INFO
Const sLocation As String = "frmProtocolInfo"


Private Sub cboProtocols_Click()
On Error GoTo VB_Error

    With WSAPROTOCOL_INFO(cboProtocols.ListIndex)
        lvwServiceFlags1.ListItems(1).SubItems(1) = CBool(.dwServiceFlags1 And XP1_CONNECTIONLESS)
        lvwServiceFlags1.ListItems(2).SubItems(1) = CBool(.dwServiceFlags1 And XP1_GUARANTEED_DELIVERY)
        lvwServiceFlags1.ListItems(3).SubItems(1) = CBool(.dwServiceFlags1 And XP1_GUARANTEED_ORDER)
        lvwServiceFlags1.ListItems(4).SubItems(1) = CBool(.dwServiceFlags1 And XP1_MESSAGE_ORIENTED)
        lvwServiceFlags1.ListItems(5).SubItems(1) = CBool(.dwServiceFlags1 And XP1_PSEUDO_STREAM)
        lvwServiceFlags1.ListItems(6).SubItems(1) = CBool(.dwServiceFlags1 And XP1_GRACEFUL_CLOSE)
        lvwServiceFlags1.ListItems(7).SubItems(1) = CBool(.dwServiceFlags1 And XP1_EXPEDITED_DATA)
        lvwServiceFlags1.ListItems(8).SubItems(1) = CBool(.dwServiceFlags1 And XP1_CONNECT_DATA)
        lvwServiceFlags1.ListItems(9).SubItems(1) = CBool(.dwServiceFlags1 And XP1_DISCONNECT_DATA)
        lvwServiceFlags1.ListItems(10).SubItems(1) = CBool(.dwServiceFlags1 And XP1_INTERRUPT)
        lvwServiceFlags1.ListItems(11).SubItems(1) = CBool(.dwServiceFlags1 And XP1_SUPPORT_BROADCAST)
        lvwServiceFlags1.ListItems(12).SubItems(1) = CBool(.dwServiceFlags1 And XP1_SUPPORT_MULTIPOINT)
        lvwServiceFlags1.ListItems(13).SubItems(1) = CBool(.dwServiceFlags1 And XP1_MULTIPOINT_CONTROL_PLANE)
        lvwServiceFlags1.ListItems(14).SubItems(1) = CBool(.dwServiceFlags1 And XP1_MULTIPOINT_DATA_PLANE)
        lvwServiceFlags1.ListItems(15).SubItems(1) = CBool(.dwServiceFlags1 And XP1_QOS_SUPPORTED)
        lvwServiceFlags1.ListItems(16).SubItems(1) = CBool(.dwServiceFlags1 And XP1_UNI_SEND)
        lvwServiceFlags1.ListItems(17).SubItems(1) = CBool(.dwServiceFlags1 And XP1_UNI_RECV)
        lvwServiceFlags1.ListItems(18).SubItems(1) = CBool(.dwServiceFlags1 And XP1_IFS_HANDLES)
        lvwServiceFlags1.ListItems(19).SubItems(1) = CBool(.dwServiceFlags1 And XP1_PARTIAL_MESSAGE)
        
        
        txtServiceFlags2.Text = int32_uint32(.dwServiceFlags2)
        txtServiceFlags3.Text = int32_uint32(.dwServiceFlags3)
        txtServiceFlags4.Text = int32_uint32(.dwServiceFlags4)
        
        txtProviderID.Text = GUID_String(.ProviderId)
        txtCatalogEntryID.Text = int32_uint32(.dwCatalogEntryId)
        txtChainLength.Text = .ProtocolChain.ChainLen
        
        lstChainEntries.Clear
        Dim lIncrement As Long
        For lIncrement = 0 To 6
            lstChainEntries.AddItem .ProtocolChain.ChainEntries(lIncrement)
        Next lIncrement
        
        chkMultipleEntries.value = IIf(.dwProviderFlags And PFL_MULTIPLE_PROTO_ENTRIES, 1, 0)
        chkRecommendedEntry.value = IIf(.dwProviderFlags And PFL_RECOMMENDED_PROTO_ENTRY, 1, 0)
        chkHidden.value = IIf(.dwProviderFlags And PFL_HIDDEN, 1, 0)
        chkMatchesProtocol0.value = IIf(.dwProviderFlags And PFL_MATCHES_PROTOCOL_ZERO, 1, 0)
        
        txtVersion.Text = .iVersion
        
        Select Case .iAddressFamily
            Case AF_UNSPEC: txtAddressFamily.Text = "Unspecified"
            Case AF_UNIX: txtAddressFamily.Text = "UNIX"
            Case AF_INET: txtAddressFamily.Text = "INET"
            Case AF_IMPLINK: txtAddressFamily.Text = "IMPLINK"
            Case AF_PUP: txtAddressFamily.Text = "PUP"
            Case AF_CHAOS: txtAddressFamily.Text = "CHAOS"
            Case AF_NS: txtAddressFamily.Text = "NS/IPX"
            Case AF_ISO: txtAddressFamily.Text = "ISO/OSI"
            Case AF_ECMA: txtAddressFamily.Text = "ECMA"
            Case AF_DATAKIT: txtAddressFamily.Text = "DATAKIT"
            Case AF_CCITT: txtAddressFamily.Text = "CCITT"
            Case AF_SNA: txtAddressFamily.Text = "SNA"
            Case AF_DECnet: txtAddressFamily.Text = "DECnet"
            Case AF_DLI: txtAddressFamily.Text = "DLI"
            Case AF_LAT: txtAddressFamily.Text = "LAT"
            Case AF_HYLINK: txtAddressFamily.Text = "HYLINK"
            Case AF_APPLETALK: txtAddressFamily.Text = "APPLETALK"
            Case AF_NETBIOS: txtAddressFamily.Text = "NETBIOS"
            Case AF_VOICEVIEW: txtAddressFamily.Text = "VOICEVIEW"
            Case AF_FIREFOX: txtAddressFamily.Text = "FIREFOX"
            Case AF_UNKNOWN1: txtAddressFamily.Text = "UNKNOWN1"
            Case AF_BAN: txtAddressFamily.Text = "BAN"
            Case AF_ATM: txtAddressFamily.Text = "ATM"
            Case AF_INET6: txtAddressFamily.Text = "INET6"
            Case AF_CLUSTER: txtAddressFamily.Text = "CLUSTER"
            Case AF_12844: txtAddressFamily.Text = "12844"
            Case AF_IRDA: txtAddressFamily.Text = "IRDA"
            Case AF_NETDES: txtAddressFamily.Text = "NETDES"
            Case AF_TCNPROCESS: txtAddressFamily.Text = "TCNPROCESS"
            Case AF_TCNMESSAGE: txtAddressFamily.Text = "TCNMESSAGE"
            Case AF_ICLFXBM: txtAddressFamily.Text = "ICLFXBM"
            Case Else: txtAddressFamily.Text = "Unknown " & .iAddressFamily
        End Select
        
        txtMaxSocketAddress.Text = .iMaxSockAddr
        txtMinSocketAddress.Text = .iMinSockAddr
        
        Select Case .iSocketType
            Case SOCK_STREAM: txtSocketType.Text = "STREAM"
            Case SOCK_DGRAM: txtSocketType.Text = "DGRAM"
            Case SOCK_RAW: txtSocketType.Text = "RAW"
            Case SOCK_RDM: txtSocketType.Text = "RDM"
            Case SOCK_SEQPACKET: txtSocketType.Text = "SEQPACKET"
            Case Else: txtSocketType.Text = "Unknown " & .iSocketType
        End Select
        
        Select Case .iProtocol
            Case IPPROTO_IP: txtProtocol.Text = "IP"
            Case IPPROTO_ICMP: txtProtocol.Text = "ICMP"
            Case IPPROTO_IGMP: txtProtocol.Text = "IGMP"
            Case IPPROTO_GGP: txtProtocol.Text = "GGP"
            Case IPPROTO_TCP: txtProtocol.Text = "TCP"
            Case IPPROTO_PUP: txtProtocol.Text = "PUP"
            Case IPPROTO_UDP: txtProtocol.Text = "UDP"
            Case IPPROTO_IDP: txtProtocol.Text = "IDP"
            Case IPPROTO_ND: txtProtocol.Text = "ND"
            Case IPPROTO_RAW: txtProtocol.Text = "RAW"
            Case Else: txtProtocol.Text = "Unknown " & .iProtocol
        End Select
        
        txtMaxProtocolOffset.Text = .iProtocolMaxOffset
        
        Select Case .iNetworkByteOrder
            Case BIGENDIAN: txtNetworkByteOrder.Text = "BIGENDIAN"
            Case LITTLEENDIAN: txtNetworkByteOrder.Text = "LITTLEENDIAN"
            Case Else: txtNetworkByteOrder.Text = "Unknown " & .iNetworkByteOrder
        End Select
        
        If .iSecurityScheme = SECURITY_PROTOCOL_NONE Then
            txtSecurityScheme.Text = "None"
        Else
            txtSecurityScheme.Text = .iSecurityScheme
        End If
        
        txtMaxMessageSize.Text = FormatNumber$(int32_uint32(.dwMessageSize), 0, , , True)
        txtReserved.Text = int32_uint32(.dwProviderReserved)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboProtocols_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwServiceFlags1
        .ColumnHeaders.Add , , "Flags"
        .ColumnHeaders.Add , , "Value"
        .ListItems.Add , , "Connectionless"
        .ListItems.Add , , "Guaranteed Delivery"
        .ListItems.Add , , "Guaranteed Order"
        .ListItems.Add , , "Message Oriented"
        .ListItems.Add , , "Psuedo Stream"
        .ListItems.Add , , "Graceful Close"
        .ListItems.Add , , "Expedited Data"
        .ListItems.Add , , "Connect Data"
        .ListItems.Add , , "Disconnect Data"
        .ListItems.Add , , "Interrupt"
        .ListItems.Add , , "Support Broadcast"
        .ListItems.Add , , "Support Multipoint"
        .ListItems.Add , , "Multipoint Control Plane"
        .ListItems.Add , , "Multipoint Data Plane"
        .ListItems.Add , , "QOS Supported"
        .ListItems.Add , , "Unidirectional Send"
        .ListItems.Add , , "Unidirectional Recieve"
        .ListItems.Add , , "IFS Handles"
        .ListItems.Add , , "Partial Message"
    End With
        
        
    If bWinsock = True Then
        Dim abBuffer() As Byte
        Dim lBufferLength As Long
        ReDim WSAPROTOCOL_INFO(0)
        
        If WSAEnumProtocols(0&, 0&, lBufferLength) = SOCKET_ERROR Then
            If Err.LastDllError <> WSAENOBUFS Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "WSAEnumProtocols")
            If (lBufferLength Mod Len(WSAPROTOCOL_INFO(0))) <> 0 Then Exit Sub
            If lBufferLength = 0 Then Exit Sub
            
            ReDim abBuffer(lBufferLength - 1)
            If WSAEnumProtocols(0&, abBuffer(0), lBufferLength) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "WSAEnumProtocols")
            
            Dim lBufferPos As Long
            Do
                Call MoveMemory(WSAPROTOCOL_INFO(UBound(WSAPROTOCOL_INFO)), abBuffer(lBufferPos), Len(WSAPROTOCOL_INFO(0)))
           
                lBufferPos = lBufferPos + Len(WSAPROTOCOL_INFO(0))
                cboProtocols.AddItem WSAPROTOCOL_INFO(UBound(WSAPROTOCOL_INFO)).szProtocol
                
                If lBufferPos < lBufferLength Then
                    ReDim Preserve WSAPROTOCOL_INFO(UBound(WSAPROTOCOL_INFO) + 1)
                Else
                    Exit Do
                End If
                
                If bShutdown = True Then Exit Do
            Loop
        End If
        
        If cboProtocols.ListCount > 0 Then cboProtocols.ListIndex = 0
    Else
        lblProtocols.Enabled = False
        cboProtocols.Enabled = False
        lblServiceFlags1.Enabled = False
        lvwServiceFlags1.Enabled = False
        lblServiceFlags2.Enabled = False
        lblServiceFlags3.Enabled = False
        lblServiceFlags4.Enabled = False
        lblProviderID.Enabled = False
        lblCatalogEntryID.Enabled = False
        lblChainLength.Enabled = False
        lblChainEntries.Enabled = False
        lstChainEntries.Enabled = False
        lblMultipleEntries.Enabled = False
        lblRecommendedEntry.Enabled = False
        lblHidden.Enabled = False
        lblMatchesProtocol0.Enabled = False
        lblVersion.Enabled = False
        lblAddressFamily.Enabled = False
        lblMaxSocketAddress.Enabled = False
        lblMinSocketAddress.Enabled = False
        lblSocketType.Enabled = False
        lblProtocol.Enabled = False
        lblMaxProtocolOffset.Enabled = False
        lblNetworkByteOrder.Enabled = False
        lblSecurityScheme.Enabled = False
        lblMaxMessageSize.Enabled = False
        lblReserved.Enabled = False
    End If
        
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
