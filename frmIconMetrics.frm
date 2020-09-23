VERSION 5.00
Begin VB.Form frmIconMetrics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Metrics"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   Icon            =   "frmIconMetrics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPIX6 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      TabIndex        =   17
      Text            =   "PIX"
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtPIX5 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Text            =   "PIX"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtPIX4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Text            =   "PIX"
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtPIX3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      TabIndex        =   8
      Text            =   "PIX"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtPIX2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Text            =   "PIX"
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtPIX1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "PIX"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtSpacingHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtSmallHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtSmallWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtSpacingWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtDefaultHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtDefaultWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblSmallHeight 
      Caption         =   "Small Icon Height"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblSmallWidth 
      Caption         =   "Small Icon Width"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblSpacingWidth 
      Caption         =   "Spacing Width"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblSpacingHeight 
      Caption         =   "Spacing Height"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblDefaultHeight 
      Caption         =   "Default Height"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblDefaultWidth 
      Caption         =   "Default Width"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmIconMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmIconMetrics"


Private Sub Form_Load()
On Error GoTo VB_Error
    
    txtDefaultWidth.Text = GetSystemMetrics(SM_CXICON)
    txtDefaultHeight.Text = GetSystemMetrics(SM_CYICON)
    txtSmallWidth.Text = GetSystemMetrics(SM_CXSMICON)
    txtSmallHeight.Text = GetSystemMetrics(SM_CYSMICON)
    txtSpacingWidth.Text = GetSystemMetrics(SM_CXICONSPACING)
    txtSpacingHeight.Text = GetSystemMetrics(SM_CYICONSPACING)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
