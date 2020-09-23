VERSION 5.00
Begin VB.Form frmErrorLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Log"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmErrorLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmErrorLog"


Private Sub Form_Load()
On Error GoTo VB_Error
    
    Forms_Loaded.bErrorLog = True
    txtLog.Text = sErrorLog
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Forms_Loaded.bErrorLog = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
