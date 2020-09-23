Attribute VB_Name = "ipifcons"
Option Explicit


Public Const MIB_IF_ADMIN_STATUS_UP As Long = 1
Public Const MIB_IF_ADMIN_STATUS_DOWN As Long = 2
Public Const MIB_IF_ADMIN_STATUS_TESTING As Long = 3

Public Const MIB_IF_OPER_STATUS_NON_OPERATIONAL As Long = 0
Public Const MIB_IF_OPER_STATUS_UNREACHABLE As Long = 1
Public Const MIB_IF_OPER_STATUS_DISCONNECTED As Long = 2
Public Const MIB_IF_OPER_STATUS_CONNECTING As Long = 3
Public Const MIB_IF_OPER_STATUS_CONNECTED As Long = 4
Public Const MIB_IF_OPER_STATUS_OPERATIONAL As Long = 5

Public Const MIB_IF_TYPE_OTHER As Long = 1
Public Const MIB_IF_TYPE_ETHERNET As Long = 6
Public Const MIB_IF_TYPE_TOKENRING As Long = 9
Public Const MIB_IF_TYPE_FDDI As Long = 15
Public Const MIB_IF_TYPE_PPP As Long = 23
Public Const MIB_IF_TYPE_LOOPBACK As Long = 24
Public Const MIB_IF_TYPE_SLIP As Long = 28

