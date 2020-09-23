Attribute VB_Name = "iprtrmib"
Option Explicit


Public Const MAXLEN_IFDESCR As Long = 256
Public Const MAXLEN_PHYSADDR As Long = 8

Public Const MIB_TCP_RTO_OTHER As Long = 1
Public Const MIB_TCP_RTO_CONSTANT As Long = 2
Public Const MIB_TCP_RTO_RSRE As Long = 3
Public Const MIB_TCP_RTO_VANJ As Long = 4

Public Const MIB_TCP_STATE_CLOSED As Long = 1
Public Const MIB_TCP_STATE_LISTEN As Long = 2
Public Const MIB_TCP_STATE_SYN_SENT As Long = 3
Public Const MIB_TCP_STATE_SYN_RCVD As Long = 4
Public Const MIB_TCP_STATE_ESTAB As Long = 5
Public Const MIB_TCP_STATE_FIN_WAIT1 As Long = 6
Public Const MIB_TCP_STATE_FIN_WAIT2 As Long = 7
Public Const MIB_TCP_STATE_CLOSE_WAIT As Long = 8
Public Const MIB_TCP_STATE_CLOSING As Long = 9
Public Const MIB_TCP_STATE_LAST_ACK As Long = 10
Public Const MIB_TCP_STATE_TIME_WAIT As Long = 11
Public Const MIB_TCP_STATE_DELETE_TCB As Long = 12


Public Type MIBICMPSTATS
    dwMsgs As Long
    dwErrors As Long
    dwDestUnreachs As Long
    dwTimeExcds As Long
    dwParmProbs As Long
    dwSrcQuenchs As Long
    dwRedirects As Long
    dwEchos As Long
    dwEchoReps As Long
    dwTimestamps As Long
    dwTimestampReps As Long
    dwAddrMasks As Long
    dwAddrMaskReps As Long
End Type

Public Type MIBICMPINFO
    icmpInStats As MIBICMPSTATS
    icmpOutStats As MIBICMPSTATS
End Type

Public Type MIB_ICMP
    stats As MIBICMPINFO
End Type

Public Type MIB_IFROW
    wszName As String * 512 'MAX_INTERFACE_NAME_LEN * 2
    dwIndex As Long
    dwType As Long
    dwMtu As Long
    dwSpeed As Long
    dwPhysAddrLen As Long
    bPhysAddr As String * MAXLEN_PHYSADDR
    dwAdminStatus As Long
    dwOperStatus As Long
    dwLastChange As Long
    dwInOctets As Long
    dwInUcastPkts As Long
    dwInNUcastPkts As Long
    dwInDiscards As Long
    dwInErrors As Long
    dwInUnknownProtos As Long
    dwOutOctets As Long
    dwOutUcastPkts As Long
    dwOutNUcastPkts As Long
    dwOutDiscards As Long
    dwOutErrors As Long
    dwOutQLen As Long
    dwDescrLen As Long
    bDescr As String * MAXLEN_IFDESCR
End Type

Public Type MIB_IFTABLE
    dwNumEntries As Long
    table(20) As MIB_IFROW
End Type

Public Type MIB_IPADDRROW
    dwAddr As Long
    dwIndex As Long
    dwMask As Long
    dwBCastAddr As Long
    dwReasmSize As Long
    unused1 As Long
    unused2 As Long
End Type

Public Type MIB_IPADDRTABLE
    dwNumEntries As Long
    table(127) As MIB_IPADDRROW
End Type

Public Type MIB_IPFORWARDROW
    dwForwardDest As Long
    dwForwardMask As Long
    dwForwardPolicy As Long
    dwForwardNextHop As Long
    dwForwardIfIndex As Long
    dwForwardType As Long
    dwForwardProto As Long
    dwForwardAge As Long
    dwForwardNextHopAS As Long
    dwForwardMetric1 As Long
    dwForwardMetric2 As Long
    dwForwardMetric3 As Long
    dwForwardMetric4 As Long
    dwForwardMetric5 As Long
End Type

Public Type MIB_IPFORWARDTABLE
    dwNumEntries As Long
    table(127) As MIB_IPFORWARDROW
End Type

Public Type MIB_IPNETROW
    dwIndex As Long
    dwPhysAddrLen As Long
    bPhysAddr As String * MAXLEN_PHYSADDR
    dwAddr As Long
    dwType As Long
End Type

Public Type MIB_IPNETTABLE
    dwNumEntries As Long
    table(127) As MIB_IPNETROW
End Type

Public Type MIB_IPSTATS
    dwForwarding As Long
    dwDefaultTTL As Long
    dwInReceives As Long
    dwInHdrErrors As Long
    dwInAddrErrors As Long
    dwForwDatagrams As Long
    dwInUnknownProtos As Long
    dwInDiscards As Long
    dwInDelivers As Long
    dwOutRequests As Long
    dwRoutingDiscards As Long
    dwOutDiscards As Long
    dwOutNoRoutes As Long
    dwReasmTimeout As Long
    dwReasmReqds As Long
    dwReasmOks As Long
    dwReasmFails As Long
    dwFragOks As Long
    dwFragFails As Long
    dwFragCreates As Long
    dwNumIf As Long
    dwNumAddr As Long
    dwNumRoutes As Long
End Type

Public Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Public Type MIB_TCPSTATS
    dwRtoAlgorithm As Long
    dwRtoMin As Long
    dwRtoMax As Long
    dwMaxConn As Long
    dwActiveOpens As Long
    dwPassiveOpens As Long
    dwAttemptFails As Long
    dwEstabResets As Long
    dwCurrEstab As Long
    dwInSegs As Long
    dwOutSegs As Long
    dwRetransSegs As Long
    dwInErrs As Long
    dwOutRsts As Long
    dwNumConns As Long
End Type

Public Type MIB_TCPTABLE
    dwNumEntries As Long
    table(127) As MIB_TCPROW
End Type

Public Type MIB_UDPROW
    dwLocalAddr As Long
    dwLocalPort As Long
End Type

Public Type MIB_UDPSTATS
    dwInDatagrams As Long
    dwNoPorts As Long
    dwInErrors As Long
    dwOutDatagrams As Long
    dwNumAddrs As Long
End Type

Public Type MIB_UDPTABLE
    dwNumEntries As Long
    table(127) As MIB_UDPROW
End Type
