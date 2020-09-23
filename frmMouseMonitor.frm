VERSION 5.00
Begin VB.Form frmMouseMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Monitor"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmMouseMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtX1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtRight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtMiddle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtWheel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtY 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtTotalClicks 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtTotalMovement 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblWheel 
      Caption         =   "Wheel"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblX1 
      Caption         =   "X"
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblTotalClicks 
      Caption         =   "Total"
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblMiddle 
      Caption         =   "Middle"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblClicks 
      Caption         =   "Clicks"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblMovement 
      Caption         =   "Movement"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lbY 
      Caption         =   "Y"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblX 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblTotalMovement 
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "frmMouseMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMouseMonitor"


Private Sub Form_Load()
On Error GoTo VB_Error
    
    Forms_Loaded.bMouseMonitor = True
    
    
    With MouseMonitor
        txtX.Text = FormatNumber(.XMovement, 0, , , True)
        txtY.Text = FormatNumber(.YMovement, 0, , , True)
        txtWheel.Text = FormatNumber(.WheelMovement, 0, , , True)
        txtTotalMovement.Text = FormatNumber(.XMovement + .YMovement, 0, , , True)
        
        txtLeft.Text = FormatNumber(.LClicks, 0, , , True)
        txtMiddle.Text = FormatNumber(.MClicks, 0, , , True)
        txtRight.Text = FormatNumber(.RClicks, 0, , , True)
        txtX1.Text = FormatNumber(.XClicks, 0, , , True)
        txtTotalClicks.Text = FormatNumber(.Clicks, 0, , , True)
    End With
    
    'InchesX = (MouseMovX * (Screen.TwipsPerPixelX / 20)) / 72
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Forms_Loaded.bMouseMonitor = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
