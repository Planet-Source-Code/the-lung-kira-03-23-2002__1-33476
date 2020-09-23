VERSION 5.00
Begin VB.Form frmAdaptersInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adapters Info"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmAdaptersInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   8175
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.CheckBox chkDHCPEnabled 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkWINS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtDHCPLeaseObtained 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtDHCPLeaseExpires 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtSecondaryWINSServerIPMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtPrimaryWINSServerIPMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtGatewayListIPMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtDHCPServerIPMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtAdapterName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtSecondaryWINSServerIPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtPrimaryWINSServerIPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtDHCPServerIPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtGatewayListIPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox cboIPAddressListIPMask 
      Height          =   315
      Left            =   2160
      TabIndex        =   22
      Top             =   3480
      Width           =   1815
   End
   Begin VB.ComboBox cboIPAddressListIPAddress 
      Height          =   315
      Left            =   2160
      TabIndex        =   20
      Top             =   3120
      Width           =   1815
   End
   Begin VB.ComboBox cboAdapters 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblDHCPLeaseExpires 
      Caption         =   "DHCP Lease Expires"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblDHCPLeaseObtained 
      Caption         =   "DHCP Lease Obtained"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblSecondaryWINSServer 
      Caption         =   "Secondary WINS Server"
      Height          =   255
      Left            =   4200
      TabIndex        =   38
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblSecondaryWINSServerIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   4200
      TabIndex        =   39
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblSecondaryWINSServerIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   4200
      TabIndex        =   41
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblPrimaryWINSServer 
      Caption         =   "Primary WINS Server"
      Height          =   255
      Left            =   4200
      TabIndex        =   33
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblPrimaryWINSServerIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   4200
      TabIndex        =   34
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblPrimaryWINSServerIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   4200
      TabIndex        =   36
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblWINS 
      Caption         =   "WINS"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblDHCPServer 
      Caption         =   "DHCP Server"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblDHCPServerIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblDHCPServerIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblGatewayListIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblGatewayListIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblGatewayList 
      Caption         =   "Gateway List"
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblIPAddressListIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblIPAddressList 
      Caption         =   "IP Address List"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblIPAddressListIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblDHCPEnabled 
      Caption         =   "DHCP Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblType 
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblIndex 
      Caption         =   "Index"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblAdapterName 
      Caption         =   "Adapter Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblAdapters 
      Caption         =   "Adapters"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdaptersInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim IP_ADAPTER_INFO() As IP_ADAPTER_INFO
Const sLocation As String = "frmAdaptersInfo"


Private Sub cboAdapters_Click()
On Error GoTo VB_Error

    With IP_ADAPTER_INFO(cboAdapters.ListIndex)
        txtAdapterName.Text = .AdapterName
        
        If Len(.AdapterName) >= .AddressLength Then
            Dim strAddress As String
            Dim lIncrement As Long
            
            strAddress = Left$(.Address, .AddressLength)
            txtAddress.Text = vbNullString
            
            For lIncrement = 1 To .AddressLength
                txtAddress.Text = txtAddress.Text & ltoa_(Asc(Mid$(.Address, lIncrement, 1)), 16)
            Next lIncrement
        End If
        
        txtIndex.Text = int32_uint32(.Index)
        
        Select Case .Type
            Case MIB_IF_TYPE_OTHER: txtType.Text = "Other"
            Case MIB_IF_TYPE_ETHERNET: txtType.Text = "Ethernet"
            Case MIB_IF_TYPE_TOKENRING: txtType.Text = "Tokenring"
            Case MIB_IF_TYPE_FDDI: txtType.Text = "FDDI"
            Case MIB_IF_TYPE_PPP: txtType.Text = "PPP"
            Case MIB_IF_TYPE_LOOPBACK: txtType.Text = "Loopback"
            Case MIB_IF_TYPE_SLIP: txtType.Text = "Slip"
            Case Else: txtType.Text = "Unknown " & .Type
        End Select
        
        chkDHCPEnabled.value = IIf(.DhcpEnabled, 1, 0)
        
        
        cboIPAddressListIPAddress.Clear
        cboIPAddressListIPAddress.AddItem .IpAddressList.IpAddress.String
        cboIPAddressListIPAddress.ListIndex = 0
        cboIPAddressListIPMask.Clear
        cboIPAddressListIPMask.AddItem .IpAddressList.IpMask.String
        cboIPAddressListIPMask.ListIndex = 0
        
        Dim IP_ADDR_STRING As IP_ADDR_STRING
        IP_ADDR_STRING = .IpAddressList
        Do
            If IP_ADDR_STRING.Next <> 0 Then
                If IsBadReadPtr(IP_ADDR_STRING.Next, Len(IP_ADDR_STRING)) = True Then
                    Call Error_API(Err.LastDllError, sLocation & "\cboAdapters_Click", "IsBadReadPtr")
                Else
                    Call MoveMemory(IP_ADDR_STRING, IP_ADDR_STRING.Next, Len(IP_ADDR_STRING))
                End If
                
                cboIPAddressListIPAddress.AddItem IP_ADDR_STRING.IpAddress.String
                cboIPAddressListIPMask.AddItem IP_ADDR_STRING.IpMask.String
            Else
                Exit Do
            End If
            
            If bShutdown = True Then Exit Do
        Loop
        
        
        txtGatewayListIPAddress.Text = .GatewayList.IpAddress.String
        txtGatewayListIPMask.Text = .GatewayList.IpMask.String
        txtDHCPServerIPAddress.Text = .DhcpServer.IpAddress.String
        txtDHCPServerIPMask.Text = .DhcpServer.IpMask.String
        chkWINS.value = IIf(.HaveWins, 1, 0)
        txtPrimaryWINSServerIPAddress.Text = .PrimaryWinsServer.IpAddress.String
        txtPrimaryWINSServerIPMask.Text = .PrimaryWinsServer.IpMask.String
        txtSecondaryWINSServerIPAddress.Text = .SecondaryWinsServer.IpAddress.String
        txtSecondaryWINSServerIPMask.Text = .SecondaryWinsServer.IpMask.String
        txtDHCPLeaseObtained.Text = .LeaseObtained
        txtDHCPLeaseExpires.Text = .LeaseExpires
        'MsgBox DateAdd("s", .LeaseObtained, "1/1/1970 12:00:00 AM")
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboAdapters_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("iphlpapi.dll", "GetAdaptersInfo") = True Then
        Dim abBuffer() As Byte
        Dim lBufferLength As Long
        ReDim IP_ADAPTER_INFO(0)
        
        lErrors = GetAdaptersInfo(ByVal 0&, lBufferLength)
        If lErrors <> ERROR_SUCCESS Then
            If lErrors <> ERROR_BUFFER_OVERFLOW Then Call Error_API(lErrors, sLocation & "\Form_Load", "GetAdaptersInfo")
            If (lBufferLength Mod Len(IP_ADAPTER_INFO(0))) <> 0 Then Exit Sub
            
            
            ReDim abBuffer(lBufferLength - 1)
            lErrors = GetAdaptersInfo(abBuffer(0), lBufferLength): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Form_Load", "GetAdaptersInfo")
            
            Dim lBufferPos As Long
            Do
                Call MoveMemory(IP_ADAPTER_INFO(UBound(IP_ADAPTER_INFO)), abBuffer(lBufferPos), Len(IP_ADAPTER_INFO(0)))
                cboAdapters.AddItem RTrim$(Str_NullTerm_Fix(IP_ADAPTER_INFO(UBound(IP_ADAPTER_INFO)).Description))
                
                lBufferPos = lBufferPos + Len(IP_ADAPTER_INFO(0))
                
                If lBufferPos < lBufferLength Then
                    ReDim Preserve IP_ADAPTER_INFO(UBound(IP_ADAPTER_INFO) + 1)
                Else
                    Exit Do
                End If
                
                If bShutdown = True Then Exit Do
            Loop
            
            If cboAdapters.ListCount > 0 Then cboAdapters.ListIndex = 0
        End If
    Else
        lblAdapters.Enabled = False
        cboAdapters.Enabled = False
        lblAdapterName.Enabled = False
        lblAddress.Enabled = False
        lblDHCPEnabled.Enabled = False
        lblIndex.Enabled = False
        lblDHCPLeaseExpires.Enabled = False
        lblDHCPLeaseObtained.Enabled = False
        lblType.Enabled = False
        lblWINS.Enabled = False
        lblIPAddressList.Enabled = False
        lblIPAddressListIPAddress.Enabled = False
        cboIPAddressListIPAddress.Enabled = False
        lblIPAddressListIPMask.Enabled = False
        cboIPAddressListIPMask.Enabled = False
        lblDHCPServer.Enabled = False
        lblDHCPServerIPAddress.Enabled = False
        lblDHCPServerIPMask.Enabled = False
        lblGatewayList.Enabled = False
        lblGatewayListIPAddress.Enabled = False
        lblGatewayListIPMask.Enabled = False
        lblPrimaryWINSServer.Enabled = False
        lblPrimaryWINSServerIPAddress.Enabled = False
        lblPrimaryWINSServerIPMask.Enabled = False
        lblSecondaryWINSServer.Enabled = False
        lblSecondaryWINSServerIPAddress.Enabled = False
        lblSecondaryWINSServerIPMask.Enabled = False
    End If

Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
