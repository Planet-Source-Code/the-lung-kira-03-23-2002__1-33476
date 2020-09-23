VERSION 5.00
Begin VB.Form frmICMPStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ICMP Statistics"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmICMPStatistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInAddressMaskReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtInAddressMaskRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtInTimeStampReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtInTimeStampRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtInParameterProblem 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtInTTLExceeded 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtInEchoRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtInRedirection 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtInSourceQuench 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtInDestinationUnreachable 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtOutAddressMaskReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtOutAddressMaskRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtOutTimeStampReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtOutTimeStampRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtOutParameterProblem 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtOutTTLExceeded 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtOutEchoRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtOutRedirection 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtOutSourceQuench 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtOutDestinationUnreachable 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtOutEchoReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtInEchoReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtInErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtOutErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   720
      Width           =   1095
   End
   Begin VB.Timer tmrICMPStatistics 
      Enabled         =   0   'False
      Interval        =   945
      Left            =   3600
      Top             =   0
   End
   Begin VB.TextBox txtOutMessages 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtInMessages 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblType 
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lbl13 
      Caption         =   "13"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lbl14 
      Caption         =   "14"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lbl12 
      Caption         =   "12"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lbl11 
      Caption         =   "11"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lbl8 
      Caption         =   "8"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lbl5 
      Caption         =   "5"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lbl4 
      Caption         =   "4"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lbl3 
      Caption         =   "3"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lbl0 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lbl17 
      Caption         =   "17"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lbl18 
      Caption         =   "18"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lblOutTimeStampRequests 
      Caption         =   "Time-Stamp Requests"
      Height          =   255
      Left            =   4200
      TabIndex        =   58
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblOutTimeStampReplies 
      Caption         =   "Time-Stamp Replies"
      Height          =   255
      Left            =   4200
      TabIndex        =   60
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblOutTTLExceeded 
      Caption         =   "TTL Exceeded"
      Height          =   255
      Left            =   4200
      TabIndex        =   54
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblOutSourceQuench 
      Caption         =   "Source Quench"
      Height          =   255
      Left            =   4200
      TabIndex        =   48
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblOutRedirection 
      Caption         =   "Redirection"
      Height          =   255
      Left            =   4200
      TabIndex        =   50
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblOutParameterProblem 
      Caption         =   "Parameter Problem"
      Height          =   255
      Left            =   4200
      TabIndex        =   56
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblOutMessages 
      Caption         =   "Messages"
      Height          =   255
      Left            =   4200
      TabIndex        =   40
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblOutErrors 
      Caption         =   "Errors"
      Height          =   255
      Left            =   4200
      TabIndex        =   42
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblOutEchoRequests 
      Caption         =   "Echo Requests"
      Height          =   255
      Left            =   4200
      TabIndex        =   52
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblOutEchoReplies 
      Caption         =   "Echo Replies"
      Height          =   255
      Left            =   4200
      TabIndex        =   44
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblOutDestinationUnreachable 
      Caption         =   "Destination Unreachable"
      Height          =   255
      Left            =   4200
      TabIndex        =   46
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblOutAddressMaskRequests 
      Caption         =   "Address Mask Requests"
      Height          =   255
      Left            =   4200
      TabIndex        =   62
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblOutAddressMaskReplies 
      Caption         =   "Address Mask Replies"
      Height          =   255
      Left            =   4200
      TabIndex        =   64
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblInTimeStampRequests 
      Caption         =   "Time-Stamp Requests"
      Height          =   255
      Left            =   720
      TabIndex        =   31
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblInTimeStampReplies 
      Caption         =   "Time-Stamp Replies"
      Height          =   255
      Left            =   720
      TabIndex        =   33
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblInTTLExceeded 
      Caption         =   "TTL Exceeded"
      Height          =   255
      Left            =   720
      TabIndex        =   27
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblInSourceQuench 
      Caption         =   "Source Quench"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblInRedirection 
      Caption         =   "Redirection"
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblInParameterProblem 
      Caption         =   "Parameter Problem"
      Height          =   255
      Left            =   720
      TabIndex        =   29
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblInMessages 
      Caption         =   "Messages"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblInErrors 
      Caption         =   "Errors"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblInEchoRequests 
      Caption         =   "Echo Requests"
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblInEchoReplies 
      Caption         =   "Echo Replies"
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblInDestinationUnreachable 
      Caption         =   "Destination Unreachable"
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblInAddressMaskRequests 
      Caption         =   "Address Mask Requests"
      Height          =   255
      Left            =   720
      TabIndex        =   35
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblInAddressMaskReplies 
      Caption         =   "Address Mask Replies"
      Height          =   255
      Left            =   720
      TabIndex        =   37
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblOut 
      Caption         =   "Out"
      Height          =   255
      Left            =   4200
      TabIndex        =   39
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblIn 
      Caption         =   "In"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmICMPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmICMPStatistics"


Private Sub Form_Load()
On Error GoTo VB_Error

    If Function_Exist("iphlpapi.dll", "GetIcmpStatistics") = True Then
        Call tmrICMPStatistics_Timer
        tmrICMPStatistics.Enabled = True
    Else
        lblInMessages.Enabled = False
        lblInErrors.Enabled = False
        lblInDestinationUnreachable.Enabled = False
        lblInTTLExceeded.Enabled = False
        lblInParameterProblem.Enabled = False
        lblInSourceQuench.Enabled = False
        lblInRedirection.Enabled = False
        lblInEchoRequests.Enabled = False
        lblInEchoReplies.Enabled = False
        lblInTimeStampRequests.Enabled = False
        lblInTimeStampReplies.Enabled = False
        lblInAddressMaskRequests.Enabled = False
        lblInAddressMaskReplies.Enabled = False
        lblOutMessages.Enabled = False
        lblOutErrors.Enabled = False
        lblOutDestinationUnreachable.Enabled = False
        lblOutTTLExceeded.Enabled = False
        lblOutParameterProblem.Enabled = False
        lblOutSourceQuench.Enabled = False
        lblOutRedirection.Enabled = False
        lblOutEchoRequests.Enabled = False
        lblOutEchoReplies.Enabled = False
        lblOutTimeStampRequests.Enabled = False
        lblOutTimeStampReplies.Enabled = False
        lblOutAddressMaskRequests.Enabled = False
        lblOutAddressMaskReplies.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    tmrICMPStatistics.Enabled = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub tmrICMPStatistics_Timer()
On Error GoTo VB_Error

    Dim MIB_ICMP As MIB_ICMP
    
    lErrors = GetIcmpStatistics(MIB_ICMP)
    If lErrors <> NO_ERROR Then
        Call Error_API(lErrors, sLocation & "\tmrICMPStatistics_Timer", "GetIcmpStatistics")
        tmrICMPStatistics.Enabled = False
        Exit Sub
    End If
    
    
    With MIB_ICMP.stats.icmpInStats
        txtInMessages.Text = FormatNumber$(int32_uint32(.dwMsgs), 0, , , True)
        txtInErrors.Text = FormatNumber$(int32_uint32(.dwErrors), 0, , , True)
        txtInDestinationUnreachable.Text = FormatNumber$(int32_uint32(.dwDestUnreachs), 0, , , True)
        txtInTTLExceeded.Text = FormatNumber$(int32_uint32(.dwTimeExcds), 0, , , True)
        txtInParameterProblem.Text = FormatNumber$(int32_uint32(.dwParmProbs), 0, , , True)
        txtInSourceQuench.Text = FormatNumber$(int32_uint32(.dwSrcQuenchs), 0, , , True)
        txtInRedirection.Text = FormatNumber$(int32_uint32(.dwRedirects), 0, , , True)
        txtInEchoRequests.Text = FormatNumber$(int32_uint32(.dwEchos), 0, , , True)
        txtInEchoReplies.Text = FormatNumber$(int32_uint32(.dwEchoReps), 0, , , True)
        txtInTimeStampRequests.Text = FormatNumber$(int32_uint32(.dwTimestamps), 0, , , True)
        txtInTimeStampReplies.Text = FormatNumber$(int32_uint32(.dwTimestampReps), 0, , , True)
        txtInAddressMaskRequests.Text = FormatNumber$(int32_uint32(.dwAddrMasks), 0, , , True)
        txtInAddressMaskReplies.Text = FormatNumber$(int32_uint32(.dwAddrMaskReps), 0, , , True)
    End With
    
    With MIB_ICMP.stats.icmpOutStats
        txtOutMessages.Text = FormatNumber$(int32_uint32(.dwMsgs), 0, , , True)
        txtOutErrors.Text = FormatNumber$(int32_uint32(.dwErrors), 0, , , True)
        txtOutDestinationUnreachable.Text = FormatNumber$(int32_uint32(.dwDestUnreachs), 0, , , True)
        txtOutTTLExceeded.Text = FormatNumber$(int32_uint32(.dwTimeExcds), 0, , , True)
        txtOutParameterProblem.Text = FormatNumber$(int32_uint32(.dwParmProbs), 0, , , True)
        txtOutSourceQuench.Text = FormatNumber$(int32_uint32(.dwSrcQuenchs), 0, , , True)
        txtOutRedirection.Text = FormatNumber$(int32_uint32(.dwRedirects), 0, , , True)
        txtOutEchoRequests.Text = FormatNumber$(int32_uint32(.dwEchos), 0, , , True)
        txtOutEchoReplies.Text = FormatNumber$(int32_uint32(.dwEchoReps), 0, , , True)
        txtOutTimeStampRequests.Text = FormatNumber$(int32_uint32(.dwTimestamps), 0, , , True)
        txtOutTimeStampReplies.Text = FormatNumber$(int32_uint32(.dwTimestampReps), 0, , , True)
        txtOutAddressMaskRequests.Text = FormatNumber$(int32_uint32(.dwAddrMasks), 0, , , True)
        txtOutAddressMaskReplies.Text = FormatNumber$(int32_uint32(.dwAddrMaskReps), 0, , , True)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\tmrICMPStatistics_Timer")
Resume Next
End Sub
