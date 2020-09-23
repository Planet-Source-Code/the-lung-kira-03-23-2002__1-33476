VERSION 5.00
Begin VB.Form frmNetworkInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Info"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmNetworkInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEnableDns 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtHostName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtDomainName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtNodeType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtLocalHostName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ComboBox cboLocalIP 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CheckBox chkNetworkPresent 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkInetIsOffline 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtNumberOfInterfaces 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.CheckBox chkEnableRouting 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox chkEnableProxy 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   3600
      Width           =   255
   End
   Begin VB.TextBox txtScopeId 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ComboBox cboDNSServerList 
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblLocalHostName 
      Caption         =   "Local Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLocalIP 
      Caption         =   "Local IP"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblNetworkPresent 
      Caption         =   "Network Present"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblInetIsOffline 
      Caption         =   "Inet Is Offline"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblNumberOfInterfaces 
      Caption         =   "Number Of Interfaces"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblEnableDns 
      Caption         =   "DNS Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblEnableRouting 
      Caption         =   "Routing Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblEnableProxy 
      Caption         =   "ARP Proxy"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblScopeId 
      Caption         =   "DHCP Scope Name"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblNodeType 
      Caption         =   "Node Type"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblDNSServerList 
      Caption         =   "DNS Server List"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblDomainName 
      Caption         =   "Domain Name"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblHostName 
      Caption         =   "Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "frmNetworkInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmNetworkInfo"


Private Sub Form_Load()
On Error GoTo VB_Error

    chkNetworkPresent.value = IIf(Right$(Right$("00000000" & ltoa_(Asc(GetSystemMetrics(SM_NETWORK)), 2), 8), 1), 1, 0)
    
    
    Dim asIP() As String
    Dim lCount As Long
    lCount = IP_Host(ComputerName_Get, asIP())
    If lCount > -1 Then txtLocalHostName.Text = Host_IP(asIP(0))
    
    With cboLocalIP
        Dim lIncrement As Long
        For lIncrement = 0 To lCount
            If asIP(lIncrement) <> vbNullString Then
                .AddItem asIP(lIncrement)
            End If
        Next lIncrement
        
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    
    
    If Function_Exist("iphlpapi.dll", "GetNetworkParams") = True Then
        Dim abBuffer() As Byte
        Dim lBufferLength As Long
        
        lErrors = GetNetworkParams(ByVal 0&, lBufferLength)
        If lErrors <> ERROR_BUFFER_OVERFLOW Then
            If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Form_Load", "GetNetworkParams")
        End If
        
        
        ReDim abBuffer(lBufferLength - 1)
        lErrors = GetNetworkParams(abBuffer(0), lBufferLength): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Form_Load", "GetNetworkParams")
        
        Dim FIXED_INFO As FIXED_INFO
        Call MoveMemory(FIXED_INFO, abBuffer(0), lBufferLength)
        
        With FIXED_INFO
            Dim IP_ADDR_STRING As IP_ADDR_STRING
            
            txtHostName.Text = Str_NullTerm_Fix(.HostName)
            txtDomainName.Text = Str_NullTerm_Fix(.DomainName)
            
            'Call MoveMemory(IP_ADDR_STRING, .CurrentDnsServer, Len(IP_ADDR_STRING))
            'txtCurrentDNSServer.Text = IP_ADDR_STRING.IpAddress.String
            
            cboDNSServerList.AddItem .DnsServerList.IpAddress.String
            cboDNSServerList.ListIndex = 0
            
            Dim lNext As Long
            lNext = .DnsServerList.Next
            Do
                If lNext <> 0 Then
                    If IsBadReadPtr(lNext, Len(IP_ADDR_STRING)) = True Then
                        Call Error_API(Err.LastDllError, sLocation & "\frmMain_Proc", "IsBadReadPtr")
                    Else
                        Call MoveMemory(IP_ADDR_STRING, ByVal lNext, Len(IP_ADDR_STRING))
                    End If
                    
                    If lNext <> IP_ADDR_STRING.Next Then
                        lNext = IP_ADDR_STRING.Next
                        cboDNSServerList.AddItem IP_ADDR_STRING.IpAddress.String
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
            
            Select Case .NodeType
                Case BROADCAST_NODETYPE: txtNodeType.Text = "Broadcast"
                Case PEER_TO_PEER_NODETYPE: txtNodeType.Text = "Peer To Peer"
                Case MIXED_NODETYPE: txtNodeType.Text = "Mixed"
                Case HYBRID_NODETYPE: txtNodeType.Text = "Hybrid"
                Case Else: txtNodeType.Text = "Unknown " & .NodeType
            End Select
            
            txtScopeId.Text = Str_NullTerm_Fix(.ScopeId)
            
            chkEnableDns.value = IIf(.EnableDns, 1, 0)
            chkEnableProxy.value = IIf(.EnableProxy, 1, 0)
            chkEnableRouting.value = IIf(.EnableRouting, 1, 0)
        End With
    Else
        lblDNSServerList.Enabled = False
        cboDNSServerList.Enabled = False
        lblDomainName.Enabled = False
        lblEnableDns.Enabled = False
        lblEnableProxy.Enabled = False
        lblEnableRouting.Enabled = False
        lblHostName.Enabled = False
        lblNodeType.Enabled = False
        lblScopeId.Enabled = False
    End If
    
    If Function_Exist("iphlpapi.dll", "GetNumberOfInterfaces") = True Then
        Dim lInterfaces As Long
        lErrors = GetNumberOfInterfaces(lInterfaces): If lErrors <> 0 Then Call Error_API(lErrors, sLocation & "\Form_Load", "GetNumberOfInterfaces")
        txtNumberOfInterfaces.Text = FormatNumber(int32_uint32(lInterfaces), 0, , , True)
    Else
        lblNumberOfInterfaces.Enabled = False
    End If
    If Function_Exist("url.dll", "InetIsOffline") = True Then
        chkInetIsOffline.value = InetIsOffline(0)
    Else
        lblInetIsOffline.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
