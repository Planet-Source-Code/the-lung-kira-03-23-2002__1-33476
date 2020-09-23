Attribute VB_Name = "iphlpapi"
Option Explicit


Public Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (ByRef pAdapterInfo As Any, ByRef pOutBufLen As Long) As Long
Public Declare Function GetIcmpStatistics Lib "iphlpapi.dll" (ByRef pStats As MIB_ICMP) As Long
Public Declare Function GetIfTable Lib "iphlpapi.dll" (ByRef pIfTable As MIB_IFTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function GetIpAddrTable Lib "iphlpapi.dll" (ByRef pIpAddrTable As MIB_IPADDRTABLE, ByRef pdwSize As Long, ByVal bOrder As Boolean) As Long
Public Declare Function GetIpForwardTable Lib "iphlpapi.dll" (ByRef pIpForwardTable As MIB_IPFORWARDTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function GetIpNetTable Lib "iphlpapi.dll" (ByRef pIpNetTable As MIB_IPNETTABLE, ByRef pdwSize As Long, ByVal bOrder As Boolean) As Long
Public Declare Function GetIpStatistics Lib "iphlpapi.dll" (ByRef pStats As MIB_IPSTATS) As Long
Public Declare Function GetIpStatisticsEx Lib "iphlpapi.dll" (ByRef pStats As MIB_IPSTATS, ByVal dwFamily As Long) As Long
Public Declare Function GetNetworkParams Lib "iphlpapi.dll" (ByRef pFixedInfo As Any, ByRef pOutBufLen As Long) As Long
Public Declare Function GetNumberOfInterfaces Lib "iphlpapi.dll" (ByRef pdwNumIf As Long) As Long
Public Declare Function GetTcpStatistics Lib "iphlpapi.dll" (ByRef pStats As MIB_TCPSTATS) As Long
Public Declare Function GetTcpStatisticsEx Lib "iphlpapi.dll" (ByRef pStats As MIB_TCPSTATS, ByVal dwFamily As Long) As Long
Public Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function GetUdpStatistics Lib "iphlpapi.dll" (ByRef pStats As MIB_UDPSTATS) As Long
Public Declare Function GetUdpStatisticsEx Lib "iphlpapi.dll" (ByRef pStats As MIB_UDPSTATS, ByVal dwFamily As Long) As Long
Public Declare Function GetUdpTable Lib "iphlpapi.dll" (ByRef pUdpTable As MIB_UDPTABLE, ByRef pdwSize As Long, ByVal bOrder As Boolean) As Long
Public Declare Function SetTcpEntry Lib "iphlpapi.dll" (ByRef pTcpRow As MIB_TCPROW) As Long
