VERSION 5.00
Begin VB.Form frmMouseInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Info"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmMouseInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPIX4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtPIX3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtPIX2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtPIX1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkMousePresent 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkMouseWheel 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtDragDropWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtDragDropHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSwapButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtCursorWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtCursorHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblMousePresent 
      Caption         =   "Mouse Present"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblDragDropWidth 
      Caption         =   "Drag Drop Width"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblDragDropHeight 
      Caption         =   "Drag Drop Height"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblButtons 
      Caption         =   "Buttons"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblMouseWheel 
      Caption         =   "Mouse Wheel"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblSwapButton 
      Caption         =   "Main Button"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblCursorWidth 
      Caption         =   "Cursor Width"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblCursorHeight 
      Caption         =   "Cursor Height"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmMouseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMouseInfo"


Private Sub Form_Load()
On Error GoTo VB_Error

    txtButtons.Text = GetSystemMetrics(SM_CMOUSEBUTTONS)
    txtCursorHeight.Text = GetSystemMetrics(SM_CYCURSOR)
    txtCursorWidth.Text = GetSystemMetrics(SM_CXCURSOR)
    txtDragDropHeight.Text = GetSystemMetrics(SM_CXDRAG)
    txtDragDropWidth.Text = GetSystemMetrics(SM_CYDRAG)
    chkMousePresent.value = GetSystemMetrics(SM_MOUSEPRESENT)
    
    If WinVersion(4010000, 0, True) = True Then
        chkMouseWheel.value = GetSystemMetrics(SM_MOUSEWHEELPRESENT)
    Else
        lblMouseWheel.Enabled = False
    End If
    
    txtSwapButton.Text = IIf(GetSystemMetrics(SM_SWAPBUTTON), "Right", "Left")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
