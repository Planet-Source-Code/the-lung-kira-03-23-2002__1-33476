VERSION 5.00
Begin VB.Form frmUDPStatisticsV6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UDP Statistics V6"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmUDPStatisticsV6.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUDPStatistics 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   600
   End
   Begin VB.TextBox txtInvalidPort 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtErrorsReceived 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtSentDatagrams 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtReceivedDatagrams 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtListenerTableEntries 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblListenerTableEntries 
      Caption         =   "Listener Table Entries"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblSentDatagrams 
      Caption         =   "Sent Datagrams"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblErrorsReceived 
      Caption         =   "Errors Received"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblInvalidPort 
      Caption         =   "Invalid Port"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblReceivedDatagrams 
      Caption         =   "Received Datagrams"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmUDPStatisticsV6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmUDPStatisticsV6"


Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("iphlpapi.dll", "GetUdpStatisticsEx") = True Then
        Call tmrUDPStatistics_Timer
        tmrUDPStatistics.Enabled = True
    Else
        lblReceivedDatagrams.Enabled = False
        lblInvalidPort.Enabled = False
        lblErrorsReceived.Enabled = False
        lblSentDatagrams.Enabled = False
        lblListenerTableEntries.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    tmrUDPStatistics.Enabled = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub tmrUDPStatistics_Timer()
On Error GoTo VB_Error

    Dim MIB_UDPSTATS As MIB_UDPSTATS
    
    lErrors = GetUdpStatisticsEx(MIB_UDPSTATS, AF_INET6)
    If lErrors <> NO_ERROR Then
        Call Error_API(lErrors, sLocation & "\tmrUDPStatistics_Timer", "GetUdpStatisticsEx")
        tmrUDPStatistics.Enabled = False
        Exit Sub
    End If
    
    With MIB_UDPSTATS
        txtReceivedDatagrams.Text = FormatNumber$(int32_uint32(.dwInDatagrams), 0, , , True)
        txtInvalidPort.Text = FormatNumber$(int32_uint32(.dwNoPorts), 0, , , True)
        txtErrorsReceived.Text = FormatNumber$(int32_uint32(.dwInErrors), 0, , , True)
        txtSentDatagrams.Text = FormatNumber$(int32_uint32(.dwOutDatagrams), 0, , , True)
        txtListenerTableEntries.Text = FormatNumber$(int32_uint32(.dwNumAddrs), 0, , , True)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\tmrUDPStatistics_Timer")
Resume Next
End Sub

