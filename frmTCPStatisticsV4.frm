VERSION 5.00
Begin VB.Form frmTCPStatisticsV4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP Statistics V4"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmTCPStatisticsV4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMaximumConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtCumulativeConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtListeningConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtConnectingConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer tmrTCPStatistics 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   3120
   End
   Begin VB.TextBox txtMaximumRTO 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtMinimumRTO 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtOutgoingResets 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtSegmentsRetransmitted 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtSegmentsSent 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtEstablishedConnectionsReset 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtEstablishedConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtIncomingErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtSegmentsReceived 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtFailedConnectionAttempts 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtRTOAlgorithm 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblSegmentsSent 
      Caption         =   "Segments Sent"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblCumulativeConnections 
      Caption         =   "Cumulative Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblOutgoingResets 
      Caption         =   "Outgoing Resets"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblIncomingErrors 
      Caption         =   "Incoming Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label lblSegmentsRetransmitted 
      Caption         =   "Segments Retransmitted"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblSegmentsReceived 
      Caption         =   "Segments Received"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblEstablishedConnections 
      Caption         =   "Established Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblEstablishedConnectionsReset 
      Caption         =   "Established Connections Reset"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblFailedConnectionAttempts 
      Caption         =   "Failed Connection Attempts"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblListeningConnections 
      Caption         =   "Listening Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblConnectingConnections 
      Caption         =   "Connecting Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblMaximumConnections 
      Caption         =   "Maximum Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblMaximumRTO 
      Caption         =   "Maximum RTO"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblMinimumRTO 
      Caption         =   "Minimum RTO"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lblRTOAlgorithm 
      Caption         =   "RTO Algorithm"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   2295
   End
End
Attribute VB_Name = "frmTCPStatisticsV4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmTCPStatisticsV4"


Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("iphlpapi.dll", "GetTcpStatistics") = True Then
        Call tmrTCPStatistics_Timer
        tmrTCPStatistics.Enabled = True
    Else
        lblRTOAlgorithm.Enabled = False
        lblMinimumRTO.Enabled = False
        lblMaximumRTO.Enabled = False
        lblMaximumConnections.Enabled = False
        lblConnectingConnections.Enabled = False
        lblListeningConnections.Enabled = False
        lblFailedConnectionAttempts.Enabled = False
        lblEstablishedConnectionsReset.Enabled = False
        lblEstablishedConnections.Enabled = False
        lblSegmentsReceived.Enabled = False
        lblSegmentsRetransmitted.Enabled = False
        lblIncomingErrors.Enabled = False
        lblOutgoingResets.Enabled = False
        lblSegmentsSent.Enabled = False
        lblCumulativeConnections.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    tmrTCPStatistics.Enabled = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub tmrTCPStatistics_Timer()
On Error GoTo VB_Error

    Dim MIB_TCPSTATS As MIB_TCPSTATS
    
    lErrors = GetTcpStatistics(MIB_TCPSTATS)
    If lErrors <> NO_ERROR Then
        Call Error_API(lErrors, sLocation & "\tmrTCPStatistics_Timer", "GetTcpStatistics")
        tmrTCPStatistics.Enabled = False
        Exit Sub
    End If
    
    With MIB_TCPSTATS
        Select Case .dwRtoAlgorithm
            Case MIB_TCP_RTO_OTHER: txtRTOAlgorithm.Text = "Other"
            Case MIB_TCP_RTO_CONSTANT: txtRTOAlgorithm.Text = "Constant Time-out"
            Case MIB_TCP_RTO_RSRE: txtRTOAlgorithm.Text = "MIL-STD-1778 Appendix B"
            Case MIB_TCP_RTO_VANJ: txtRTOAlgorithm.Text = "Van Jacobson's Algorithm"
            Case Else: txtRTOAlgorithm.Text = "Unknown " & int32_uint32(.dwRtoAlgorithm)
        End Select
        
        txtMinimumRTO.Text = FormatNumber$(int32_uint32(.dwRtoMin), 0, , , True)
        txtMaximumRTO.Text = FormatNumber$(int32_uint32(.dwRtoMax), 0, , , True)
        txtMaximumConnections.Text = FormatNumber$(int32_uint32(.dwMaxConn), 0, , , True)
        txtConnectingConnections.Text = FormatNumber$(int32_uint32(.dwActiveOpens), 0, , , True)
        txtListeningConnections.Text = FormatNumber$(int32_uint32(.dwPassiveOpens), 0, , , True)
        txtFailedConnectionAttempts.Text = FormatNumber$(int32_uint32(.dwAttemptFails), 0, , , True)
        txtEstablishedConnectionsReset.Text = FormatNumber$(int32_uint32(.dwEstabResets), 0, , , True)
        txtEstablishedConnections.Text = FormatNumber$(int32_uint32(.dwCurrEstab), 0, , , True)
        txtSegmentsReceived.Text = FormatNumber$(int32_uint32(.dwInSegs), 0, , , True)
        txtSegmentsRetransmitted.Text = FormatNumber$(int32_uint32(.dwRetransSegs), 0, , , True)
        txtIncomingErrors.Text = FormatNumber$(int32_uint32(.dwInErrs), 0, , , True)
        txtOutgoingResets.Text = FormatNumber$(int32_uint32(.dwOutRsts), 0, , , True)
        txtSegmentsSent.Text = FormatNumber$(int32_uint32(.dwOutSegs), 0, , , True)
        txtCumulativeConnections.Text = FormatNumber$(int32_uint32(.dwNumConns), 0, , , True)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\tmrTCPStatistics_Timer")
Resume Next
End Sub
