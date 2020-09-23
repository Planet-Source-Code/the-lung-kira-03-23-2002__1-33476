Attribute VB_Name = "iptypes"
Option Explicit


Public Const BROADCAST_NODETYPE As Long = 1
Public Const PEER_TO_PEER_NODETYPE As Long = 2
Public Const MIXED_NODETYPE As Long = 4
Public Const HYBRID_NODETYPE As Long = 8

Public Const MAX_ADAPTER_DESCRIPTION_LENGTH As Long = 128
Public Const MAX_ADAPTER_NAME_LENGTH As Long = 256
Public Const MAX_ADAPTER_ADDRESS_LENGTH As Long = 8
Public Const DEFAULT_MINIMUM_ENTITIES As Long = 32
Public Const MAX_HOSTNAME_LEN As Long = 128
Public Const MAX_DOMAIN_NAME_LEN As Long = 128
Public Const MAX_SCOPE_ID_LEN As Long = 256


Public Type IP_ADDRESS_STRING
    String As String * 16 '4 x 4
End Type

Public Type IP_ADDR_STRING
    Next As Long
    IpAddress As IP_ADDRESS_STRING
    IpMask As IP_ADDRESS_STRING
    Context As Long
End Type

Public Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName As String * 260 'MAX_ADAPTER_NAME_LENGTH + 4
    Description As String * 132 'MAX_ADAPTER_DESCRIPTION_LENGTH + 4
    AddressLength As Long
    Address As String * MAX_ADAPTER_ADDRESS_LENGTH
    Index As Long
    Type As Long
    DhcpEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As IP_ADDR_STRING
    GatewayList As IP_ADDR_STRING
    DhcpServer As IP_ADDR_STRING
    HaveWins As Long
    PrimaryWinsServer As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained As Long
    LeaseExpires As Long
End Type

Public Type FIXED_INFO
    HostName As String * 132            'MAX_HOSTNAME_LEN + 4
    DomainName As String * 132          'MAX_DOMAIN_NAME_LEN + 4
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId As String * 260             'MAX_SCOPE_ID_LEN + 4
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type
