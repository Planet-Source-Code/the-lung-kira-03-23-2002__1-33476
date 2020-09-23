Attribute VB_Name = "winsock2"
Option Explicit


Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Public Declare Function connect Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As Any, ByVal namelen As Long) As Long
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (ByRef addr As Long, ByRef addrlen As Long, ByVal addrType As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal port As Long, ByRef proto As Any) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Long
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef fromaddr As Any, ByRef fromlen As Long) As Long
Public Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef toaddr As Any, ByVal tolen As Long) As Long
Public Declare Function shutdown Lib "ws2_32.dll" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAEnumProtocols Lib "ws2_32.dll" Alias "WSAEnumProtocolsA" (ByVal lpiProtocols As Long, ByRef lpProtocolBuffer As Any, ByRef lpdwBufferLength As Long) As Long
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long


Public Const INVALID_SOCKET As Long = &HFFFF
Public Const SOCKET_ERROR As Long = -1
Public Const BIGENDIAN As Long = &H0
Public Const LITTLEENDIAN As Long = &H1

Public Const AF_UNSPEC As Long = 0
Public Const AF_UNIX As Long = 1
Public Const AF_INET As Long = 2
Public Const AF_IMPLINK As Long = 3
Public Const AF_PUP As Long = 4
Public Const AF_CHAOS As Long = 5
Public Const AF_NS As Long = 6
Public Const AF_IPX As Long = AF_NS
Public Const AF_ISO As Long = 7
Public Const AF_OSI As Long = AF_ISO
Public Const AF_ECMA As Long = 8
Public Const AF_DATAKIT As Long = 9
Public Const AF_CCITT As Long = 10
Public Const AF_SNA As Long = 11
Public Const AF_DECnet As Long = 12
Public Const AF_DLI As Long = 13
Public Const AF_LAT As Long = 14
Public Const AF_HYLINK As Long = 15
Public Const AF_APPLETALK As Long = 16
Public Const AF_NETBIOS As Long = 17
Public Const AF_VOICEVIEW As Long = 18
Public Const AF_FIREFOX As Long = 19
Public Const AF_UNKNOWN1 As Long = 20
Public Const AF_BAN As Long = 21
Public Const AF_ATM As Long = 22
Public Const AF_INET6 As Long = 23
Public Const AF_CLUSTER As Long = 24
Public Const AF_12844 As Long = 25
Public Const AF_IRDA As Long = 26
Public Const AF_NETDES As Long = 28
Public Const AF_TCNPROCESS As Long = 29
Public Const AF_TCNMESSAGE As Long = 30
Public Const AF_ICLFXBM As Long = 31

Public Const FD_READ As Long = &H1
Public Const FD_WRITE As Long = &H2
Public Const FD_OOB As Long = &H4
Public Const FD_ACCEPT As Long = &H8
Public Const FD_CONNECT As Long = &H10
Public Const FD_CLOSE As Long = &H20

Public Const INADDR_ANY As Long = &H0
Public Const INADDR_LOOPBACK As Long = &H7F000001
Public Const INADDR_BROADCAST As Long = &HFFFFFFFF
Public Const INADDR_NONE As Long = &HFFFFFFFF

Public Const IPPROTO_IP As Long = 0
Public Const IPPROTO_ICMP As Long = 1
Public Const IPPROTO_IGMP As Long = 2
Public Const IPPROTO_GGP As Long = 3
Public Const IPPROTO_TCP As Long = 6
Public Const IPPROTO_PUP As Long = 12
Public Const IPPROTO_UDP As Long = 17
Public Const IPPROTO_IDP As Long = 22
Public Const IPPROTO_ND As Long = 77
Public Const IPPROTO_RAW As Long = 255
Public Const IPPROTO_MAX As Long = 256

Public Const MAX_PROTOCOL_CHAIN As Long = 7

Public Const PFL_MULTIPLE_PROTO_ENTRIES As Long = &H1
Public Const PFL_RECOMMENDED_PROTO_ENTRY As Long = &H2
Public Const PFL_HIDDEN As Long = &H4
Public Const PFL_MATCHES_PROTOCOL_ZERO As Long = &H8

Public Const SD_RECEIVE As Long = &H0
Public Const SD_SEND As Long = &H1
Public Const SD_BOTH As Long = &H2

Public Const SECURITY_PROTOCOL_NONE As Long = &H0

Public Const SOCK_STREAM As Long = 1
Public Const SOCK_DGRAM As Long = 2
Public Const SOCK_RAW As Long = 3
Public Const SOCK_RDM As Long = 4
Public Const SOCK_SEQPACKET As Long = 5

Public Const WSA_DESCRIPTION_LEN As Long = 256
Public Const WSA_SYS_STATUS_LEN As Long = 128
Public Const WSAPROTOCOL_LEN As Long = 255

Public Const XP1_CONNECTIONLESS As Long = &H1
Public Const XP1_GUARANTEED_DELIVERY As Long = &H2
Public Const XP1_GUARANTEED_ORDER As Long = &H4
Public Const XP1_MESSAGE_ORIENTED As Long = &H8
Public Const XP1_PSEUDO_STREAM As Long = &H10
Public Const XP1_GRACEFUL_CLOSE As Long = &H20
Public Const XP1_EXPEDITED_DATA As Long = &H40
Public Const XP1_CONNECT_DATA As Long = &H80
Public Const XP1_DISCONNECT_DATA As Long = &H100
Public Const XP1_SUPPORT_BROADCAST As Long = &H200
Public Const XP1_SUPPORT_MULTIPOINT As Long = &H400
Public Const XP1_MULTIPOINT_CONTROL_PLANE As Long = &H800
Public Const XP1_MULTIPOINT_DATA_PLANE As Long = &H1000
Public Const XP1_QOS_SUPPORTED As Long = &H2000
Public Const XP1_INTERRUPT As Long = &H4000
Public Const XP1_UNI_SEND As Long = &H8000
Public Const XP1_UNI_RECV As Long = &H10000
Public Const XP1_IFS_HANDLES As Long = &H20000
Public Const XP1_PARTIAL_MESSAGE As Long = &H40000


Public Type hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public Type in_addr
    S_un As Long
End Type

Public Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type

Public Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As in_addr
    sin_zero As String * 8
End Type

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257   'WSADESCRIPTION_LEN + 1
    szSystemStatus As String * 129  'WSASYS_STATUS_LEN + 1
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Type WSAPROTOCOLCHAIN
    ChainLen As Long
    ChainEntries(MAX_PROTOCOL_CHAIN - 1) As Long
End Type

Public Type WSAPROTOCOL_INFO
    dwServiceFlags1 As Long
    dwServiceFlags2 As Long
    dwServiceFlags3 As Long
    dwServiceFlags4 As Long
    dwProviderFlags As Long
    ProviderId As GUID
    dwCatalogEntryId As Long
    ProtocolChain As WSAPROTOCOLCHAIN
    iVersion As Long
    iAddressFamily As Long
    iMaxSockAddr As Long
    iMinSockAddr As Long
    iSocketType As Long
    iProtocol As Long
    iProtocolMaxOffset As Long
    iNetworkByteOrder As Long
    iSecurityScheme As Long
    dwMessageSize As Long
    dwProviderReserved As Long
    szProtocol As String * 256 'WSAPROTOCOL_LEN + 1
End Type
