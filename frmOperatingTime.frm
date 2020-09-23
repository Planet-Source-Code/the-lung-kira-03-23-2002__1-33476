VERSION 5.00
Begin VB.Form frmOperatingTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operating Time"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmOperatingTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpTime 
      Interval        =   945
      Left            =   1080
      Top             =   120
   End
   Begin VB.TextBox txtDays 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtHours 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtMinutes 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtSeconds 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtMilliseconds 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtUnFormatted 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblDays 
      Caption         =   "Days"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblHours 
      Caption         =   "Hours"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Minutes"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblSeconds 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblMilliseconds 
      Caption         =   "Milliseconds"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblUnFormatted 
      Caption         =   "UnFormatted"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmOperatingTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmOperatingTime"


Private Sub Form_Load()
On Error GoTo VB_Error

    tmrUpTime_Timer
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    tmrUpTime.Enabled = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub tmrUpTime_Timer()
On Error GoTo VB_Error

    Dim lUpTime As Long
    lUpTime = GetTickCount
    
    txtUnFormatted.Text = FormatNumber$(lUpTime, 0, , , True)
    txtMilliseconds.Text = (lUpTime - ((lUpTime \ 1000) * 1000))
    lUpTime = lUpTime \ 1000
    txtSeconds.Text = (lUpTime - ((lUpTime \ 60) * 60))
    lUpTime = lUpTime \ 60
    txtMinutes.Text = (lUpTime - ((lUpTime \ 60) * 60))
    lUpTime = lUpTime \ 60
    txtHours.Text = (lUpTime - ((lUpTime \ 24) * 24))
    lUpTime = lUpTime \ 24
    txtDays.Text = (lUpTime)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\tmrUpTime_Timer")
Resume Next
End Sub
