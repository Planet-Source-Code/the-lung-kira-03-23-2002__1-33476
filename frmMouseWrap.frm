VERSION 5.00
Begin VB.Form frmMouseWrap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Wrap"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmMouseWrap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWraps 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblWraps 
      Caption         =   "Number of Wrap"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMouseWrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMouseWrap"


Private Sub Form_Load()
On Error GoTo VB_Error
    
    Forms_Loaded.bMouseWrap = True
    
    txtWraps.Text = FormatNumber(MouseMonitor.Wrap, 0, , , True)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Forms_Loaded.bMouseWrap = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
