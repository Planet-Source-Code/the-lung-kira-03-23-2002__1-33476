VERSION 5.00
Begin VB.Form frmIPStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Statistics"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmIPStatistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRoutesInRoutingTable 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtIPAddressesOnComputer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtIPForwarding 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtInterfacesOnComputer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtTransmittedDatagramsDiscarded 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtOutgoingDatagramsDiscarded 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtReceivedDatagramsDiscarded 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtReceivedAddressErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtReceivedHeaderErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtDatagramsNoRoute 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtDatagramsMissingFragments 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtFailedFragmentations 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtSuccessfulFragmentations 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtDatagramsFragmented 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtFailedReassemblies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtSuccessfulReassemblies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtReceivedDatagramsDelivered 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtDatagramsForwarded 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Timer tmrIPStatistics 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   0
   End
   Begin VB.TextBox txtOutgoingDatagramsRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtDatagramsRequiringReassembly 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtDatagramsUnknownProtocol 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtDatagramsReceived 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtDefaultTTL 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblTransmittedDatagramsDiscarded 
      Caption         =   "Transmitted Datagrams Discarded"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblOutgoingDatagramsRequests 
      Caption         =   "Outgoing Datagrams Requests"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblIPForwarding 
      Caption         =   "IP Forwarding"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label lblDatagramsMissingFragments 
      Caption         =   "Datagrams Missing Fragments"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblReceivedHeaderErrors 
      Caption         =   "Received Header Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblReceivedAddressErrors 
      Caption         =   "Received Address Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblDatagramsUnknownProtocol 
      Caption         =   "Datagrams Unknown Protocol"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblInterfacesOnComputer 
      Caption         =   "Interfaces On Computer"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label lblIPAddressesOnComputer 
      Caption         =   "IP Addresses On Computer"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lblDefaultTTL 
      Caption         =   "Default TTL"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label lblDatagramsNoRoute 
      Caption         =   "Datagrams No Route"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lblRoutesInRoutingTable 
      Caption         =   "Routes In Routing Table"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label lblSuccessfulReassemblies 
      Caption         =   "Successful Reassemblies"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblFailedReassemblies 
      Caption         =   "Failed Reassemblies"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblDatagramsRequiringReassembly 
      Caption         =   "Datagrams Requiring Reassembly"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblSuccessfulFragmentations 
      Caption         =   "Successful Fragmentations"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblFailedFragmentations 
      Caption         =   "Failed Fragmentations"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblDatagramsFragmented 
      Caption         =   "Datagrams Fragmented"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblReceivedDatagramsDiscarded 
      Caption         =   "Received Datagrams Discarded"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label lblReceivedDatagramsDelivered 
      Caption         =   "Received Datagrams Delivered"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblDatagramsForwarded 
      Caption         =   "Datagrams Forwarded"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblOutgoingDatagramsDiscarded 
      Caption         =   "Outgoing Datagrams Discarded"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblDatagramsReceived 
      Caption         =   "Datagrams Received"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmIPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("iphlpapi.dll", "GetIpStatistics") = True Then
        tmrIPStatistics_Timer
        tmrIPStatistics.Enabled = True
    Else
        lblIPForwarding.Enabled = False
        lblDefaultTTL.Enabled = False
        lblDatagramsReceived.Enabled = False
        lblReceivedHeaderErrors.Enabled = False
        lblReceivedAddressErrors.Enabled = False
        lblDatagramsForwarded.Enabled = False
        lblDatagramsUnknownProtocol.Enabled = False
        lblReceivedDatagramsDiscarded.Enabled = False
        lblReceivedDatagramsDelivered.Enabled = False
        lblOutgoingDatagramsRequests.Enabled = False
        lblOutgoingDatagramsDiscarded.Enabled = False
        lblTransmittedDatagramsDiscarded.Enabled = False
        lblDatagramsNoRoute.Enabled = False
        lblDatagramsMissingFragments.Enabled = False
        lblDatagramsRequiringReassembly.Enabled = False
        lblSuccessfulReassemblies.Enabled = False
        lblFailedReassemblies.Enabled = False
        lblSuccessfulFragmentations.Enabled = False
        lblFailedFragmentations.Enabled = False
        lblDatagramsFragmented.Enabled = False
        lblInterfacesOnComputer.Enabled = False
        lblIPAddressesOnComputer.Enabled = False
        lblRoutesInRoutingTable.Enabled = False
    End If
    
Exit Sub
VB_Error:
VB_Errors
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    tmrIPStatistics.Enabled = False
    
Exit Sub
VB_Error:
VB_Errors
End Sub

Private Sub tmrIPStatistics_Timer()
On Error GoTo VB_Error

    Dim MIB_IPSTATS As MIB_IPSTATS
    
    If Not GetIpStatistics(MIB_IPSTATS) = NO_ERROR Then Errors Err.LastDllError, "GetIpStatistics"
    
    With MIB_IPSTATS
        txtIPForwarding.Text = CBool(.dwForwarding)
        txtDefaultTTL.Text = Format_Comma(int32_uint32(.dwDefaultTTL))
        txtDatagramsReceived.Text = Format_Comma(int32_uint32(.dwInReceives))
        txtReceivedHeaderErrors.Text = Format_Comma(int32_uint32(.dwInHdrErrors))
        txtReceivedAddressErrors.Text = Format_Comma(int32_uint32(.dwInAddrErrors))
        txtDatagramsForwarded.Text = Format_Comma(int32_uint32(.dwForwDatagrams))
        txtDatagramsUnknownProtocol.Text = Format_Comma(int32_uint32(.dwInUnknownProtos))
        txtReceivedDatagramsDiscarded.Text = Format_Comma(int32_uint32(.dwInDiscards))
        txtReceivedDatagramsDelivered.Text = Format_Comma(int32_uint32(.dwInDelivers))
        txtOutgoingDatagramsRequests.Text = Format_Comma(int32_uint32(.dwOutRequests))
        txtOutgoingDatagramsDiscarded.Text = Format_Comma(int32_uint32(.dwRoutingDiscards))
        txtTransmittedDatagramsDiscarded.Text = Format_Comma(int32_uint32(.dwOutDiscards))
        txtDatagramsNoRoute.Text = Format_Comma(int32_uint32(.dwOutNoRoutes))
        txtDatagramsMissingFragments.Text = Format_Comma(int32_uint32(.dwReasmTimeout))
        txtDatagramsRequiringReassembly.Text = Format_Comma(int32_uint32(.dwReasmReqds))
        txtSuccessfulReassemblies.Text = Format_Comma(int32_uint32(.dwReasmOks))
        txtFailedReassemblies.Text = Format_Comma(int32_uint32(.dwReasmFails))
        txtSuccessfulFragmentations.Text = Format_Comma(int32_uint32(.dwFragOks))
        txtFailedFragmentations.Text = Format_Comma(int32_uint32(.dwFragFails))
        txtDatagramsFragmented.Text = Format_Comma(int32_uint32(.dwFragCreates))
        txtInterfacesOnComputer.Text = Format_Comma(int32_uint32(.dwNumIf))
        txtIPAddressesOnComputer.Text = Format_Comma(int32_uint32(.dwNumAddr))
        txtRoutesInRoutingTable.Text = Format_Comma(int32_uint32(.dwNumRoutes))
    End With
    
Exit Sub
VB_Error:
VB_Errors
End Sub
