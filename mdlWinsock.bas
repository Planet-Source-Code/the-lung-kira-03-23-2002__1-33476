Attribute VB_Name = "mdlWinsock"
Option Explicit


Public Type ICMPHDR4
    Type As Byte
    Code As Byte
    checksum As Integer
    ID As Integer
    Seq As Integer
    data As String
End Type

Public Type IPHDR4
    VIHL As Byte
    TOS  As Byte
    TotLen As Integer
    ID As Integer
    FlagOff As Integer
    TTL As Byte
    protocol As Byte
    checksum As Integer
    iaSrc As Long
    iaDst As Long
    options As String
    data As String
End Type

Public Type ICMP_Packet4
    IPHDR4 As IPHDR4
    ICMPHDR4 As ICMPHDR4
End Type


Public Const ICMP_ECHO_REPLY As Long = 0
Public Const ICMP_DESTINATION_UNREACHABLE As Long = 3
Public Const ICMP_SOURCE_QUENCH As Long = 4
Public Const ICMP_REDIRECT As Long = 5
Public Const ICMP_ECHO As Long = 8
Public Const ICMP_ROUTER_ADVERTISEMENT As Long = 9
Public Const ICMP_ROUTER_SELECTION As Long = 10
Public Const ICMP_TIME_EXCEEDED As Long = 11
Public Const ICMP_PARAMETER_PROBLEM As Long = 12
Public Const ICMP_TIMESTAMP As Long = 13
Public Const ICMP_TIMESTAMP_REPLY As Long = 14
Public Const ICMP_INFORMATION_REQUEST As Long = 15
Public Const ICMP_INFORMATION_REPLY As Long = 16
Public Const ICMP_ADDRESS_MASK_REQUEST As Long = 17
Public Const ICMP_ADDRESS_MASK_REPLY As Long = 18
Public Const ICMP_ADDRESS_TRACEROUTE As Long = 30
Public Const ICMP_DATAGRAM_CONVERSION_ERROR As Long = 31

Const sLocation As String = "mdlWinsock"


Public Function Host_IP(ByVal sIP As String) As String
On Error GoTo VB_Error

    If bWinsock = True Then
        'If Function_Exist("ws2_32.dll", "getnameinfo") = True Then
        
        'Else
            Dim hostent As hostent
            Dim lIP As Long
            Dim sHost As String * 255
            
            lIP = inet_addr(sIP)
            
            Dim lPtr As Long
            lPtr = gethostbyaddr(lIP, Len(lIP), AF_INET)
            If lPtr = 0 Then
                Call Error_API(Err.LastDllError, sLocation & "\Host_IP", "gethostbyaddr")
                Exit Function
            End If
            
            If IsBadReadPtr(lPtr, Len(hostent)) = True Then
                Call Error_API(Err.LastDllError, sLocation & "\Host_IP", "IsBadReadPtr")
            Else
                Call MoveMemory(hostent, ByVal lPtr, Len(hostent))
            End If
            If IsBadReadPtr(hostent.h_name, 255) = True Then
                Call Error_API(Err.LastDllError, sLocation & "\Host_IP", "IsBadReadPtr")
            Else
                Call MoveMemory(ByVal sHost, ByVal hostent.h_name, 255)
            End If
            
            Host_IP = Str_NullTerm_Fix(sHost)
        'End If
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\IP_String")
Resume Next
End Function

Public Function IP_Host(ByVal sHost As String, ByRef asIP() As String) As Long
On Error GoTo VB_Error
    
    If bWinsock = True Then
        'If Function_Exist("ws2_32.dll", "getaddrinfo") = True Then
        
        'Else
            Dim hostent As hostent
            Dim lHostIp As Long
            Dim sIP As String
            Dim lIP As Long
            Dim lValue As Long
            Dim lCount As Long
            
            ReDim asIP(lCount)
            
            Dim lPtr As Long
            lPtr = gethostbyname(sHost)
            If lPtr = 0 Then
                Call Error_API(Err.LastDllError, sLocation & "\IP_Host", "gethostbyname")
                lCount = -1
                Exit Function
            End If
            
            If IsBadReadPtr(lPtr, Len(hostent)) = True Then
                Call Error_API(Err.LastDllError, sLocation & "\IP_Host", "IsBadReadPtr")
            Else
                Call MoveMemory(hostent, ByVal lPtr, Len(hostent))
            End If
            If IsBadReadPtr(hostent.h_addr_list, hostent.h_length) = True Then
                Call Error_API(Err.LastDllError, sLocation & "\IP_Host", "IsBadReadPtr")
            Else
                Call MoveMemory(lHostIp, ByVal hostent.h_addr_list, hostent.h_length)
            End If
            If lHostIp = 0 Then Exit Function
            
            Do
                If IsBadReadPtr(lHostIp, hostent.h_length) = True Then
                    Call Error_API(Err.LastDllError, sLocation & "\IP_Host", "IsBadReadPtr")
                Else
                    Call MoveMemory(lValue, ByVal lHostIp, hostent.h_length)
                End If
                
                sIP = IP_String(lValue)
                asIP(lCount) = sIP
                
                
                hostent.h_addr_list = hostent.h_addr_list + Len(hostent.h_addr_list)
                
                If IsBadReadPtr(hostent.h_addr_list, hostent.h_length) = True Then
                    Call Error_API(Err.LastDllError, sLocation & "\IP_Host", "IsBadReadPtr")
                Else
                    Call MoveMemory(lHostIp, ByVal hostent.h_addr_list, hostent.h_length)
                End If
                
                If lHostIp = 0 Then
                    Exit Do
                Else
                    lCount = lCount + 1
                    ReDim Preserve asIP(lCount)
                End If
                
                If bShutdown = True Then Exit Do
            Loop
            
            IP_Host = lCount
        'End If
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\IP_Host")
Resume Next
End Function

Public Function IP_String(ByVal lIP As Long) As String
On Error GoTo VB_Error
    
    If bWinsock = True Then
        'If Function_Exist("ws2_32.dll", "getaddrinfo") = True Then
        
        'Else
            Dim sIP As String
            
            lIP = inet_ntoa(lIP)
            sIP = String$(lstrlen(lIP), 0) '15
            
            If IsBadReadPtr(lIP, Len(sIP)) = True Then
                Call Error_API(Err.LastDllError, sLocation & "\IP_String", "IsBadReadPtr")
            Else
                Call MoveMemory(ByVal sIP, ByVal lIP, Len(sIP))
            End If
            
            IP_String = Str_NullTerm_Fix(sIP)
        'End If
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\IP_String")
Resume Next
End Function

Public Sub Socket_Close(ByRef lSocket As Long)
On Error GoTo VB_Error
    
    If bWinsock = True Then
        If lSocket <> 0 Then
            If closesocket(lSocket) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\Socket_Close", "closesocket")
            lSocket = 0
        End If
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Socket_Close")
Resume Next
End Sub

Public Sub WSv4_Start(ByVal sHostIP As String, ByVal lPort As Long, ByVal hwnd As Long, ByVal lType As Long, ByRef lSocket As Long, ByRef sockaddr As sockaddr_in)
On Error GoTo VB_Error
    
    
    Call Socket_Close(lSocket)
    
    
    Dim lAddr As Long
    lAddr = inet_addr(sHostIP & vbNullChar)
    
    If lAddr = INADDR_NONE Then
        Dim asIP() As String
        Dim lCount As Long
        lCount = IP_Host(sHostIP & vbNullChar, asIP())
        
        If lCount > -1 Then
            If asIP(0) <> vbNullString Then
                lAddr = inet_addr(asIP(0) & vbNullChar)
                If lAddr = INADDR_NONE Then Call Error_API(Err.LastDllError, sLocation & "\WSv4_Start", "inet_addr")
            End If
        End If
    End If
        
        
    With sockaddr
        .sin_addr.S_un = lAddr
        .sin_family = AF_INET
        .sin_port = uint16_int16(htons(lPort))
        .sin_zero = String$(8, 0)
    End With
    
    
    Select Case lType
        Case 0
            lSocket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP): If lSocket = INVALID_SOCKET Then Call Error_API(Err.LastDllError, sLocation & "\WSv4_Start", "socket")
            If WSAAsyncSelect(lSocket, hwnd, WM_WINSOCK_MSG, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\WSv4_Start", "WSAAsyncSelect")
        Case 1
            lSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP): If lSocket = INVALID_SOCKET Then Call Error_API(Err.LastDllError, sLocation & "\WSv4_Start", "socket")
            If connect(lSocket, sockaddr, Len(sockaddr)) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\WSv4_Start", "connect")
            If WSAAsyncSelect(lSocket, hwnd, WM_WINSOCK_MSG, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\WSv4_Start", "WSAAsyncSelect")
    End Select
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\WSv4_Start")
Resume Next
End Sub

Public Sub Winsock_Start()
On Error GoTo VB_Error

    If Function_Exist("ws2_32.dll", "WSAStartup") = True Then
        Dim WSADATA As WSADATA
        lErrors = WSAStartup(MAKEWORD(2, 2), WSADATA)
        If lErrors <> 0 Then
            Call Error_API(lErrors, sLocation & "\Winsock_Start", "WSAStartup")
            bWinsock = False
        Else
            bWinsock = True
            wsBuffer = String$(65536, 0)
        End If
    Else
        bWinsock = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Winsock_Start")
Resume Next
End Sub

Public Sub Winsock_Stop()
On Error GoTo VB_Error
    
    If bWinsock = True Then
        lErrors = WSACleanup: If lErrors = SOCKET_ERROR Then Call Error_API(lErrors, sLocation & "\Winsock_Stop", "WSACleanup")
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Winsock_Stop")
Resume Next
End Sub
