VERSION 5.00
Begin VB.Form frmExtra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmExtra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmExtra.frx":000C
   ScaleHeight     =   8520
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblExeVersionText 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label lblEmailText 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "lung@vcn.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblPageText 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.lung.vcn.com/"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label lblExeVersion 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Kira.Exe Version"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label lblPage 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Page"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblEMail 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmExtra"


Private Sub Form_Load()
On Error GoTo VB_Error
    
    lblExeVersionText.Caption = sAppVer
        
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lblEMail_Click")
Resume Next
End Sub

Private Sub lblEMail_Click()
On Error GoTo VB_Error

    Dim SHELLEXECUTEINFO As SHELLEXECUTEINFO
    With SHELLEXECUTEINFO
        .cbSize = Len(SHELLEXECUTEINFO)
        .fMask = SEE_MASK_FLAG_NO_UI
        .hwnd = frmExtra.hwnd
        .lpVerb = "open"
        .lpFile = "mailto:" & lblEmailText.Caption
        .nShow = SW_SHOW
    End With
    
    If ShellExecuteEx(SHELLEXECUTEINFO) = False Then Call Error_API(Err.LastDllError, sLocation & "\lblEMail_Click", "ShellExecuteEx")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lblEMail_Click")
Resume Next
End Sub

Private Sub lblHelp_Click()
On Error GoTo VB_Error

    Dim SHELLEXECUTEINFO As SHELLEXECUTEINFO
    With SHELLEXECUTEINFO
        .cbSize = Len(SHELLEXECUTEINFO)
        .fMask = SEE_MASK_FLAG_NO_UI
        .hwnd = frmExtra.hwnd
        .lpVerb = "open"
        .lpFile = sAppPath & "\Kira.chm"
        .nShow = SW_SHOW
    End With
    
    If ShellExecuteEx(SHELLEXECUTEINFO) = False Then Call Error_API(Err.LastDllError, sLocation & "\lblHelp_Click", "ShellExecuteEx")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lblHelp_Click")
Resume Next
End Sub

Private Sub lblPage_Click()
On Error GoTo VB_Error

    Dim SHELLEXECUTEINFO As SHELLEXECUTEINFO
    With SHELLEXECUTEINFO
        .cbSize = Len(SHELLEXECUTEINFO)
        .fMask = SEE_MASK_FLAG_NO_UI
        .hwnd = frmExtra.hwnd
        .lpVerb = "open"
        .lpFile = lblPageText.Caption
        .nShow = SW_SHOW
    End With
    
    If ShellExecuteEx(SHELLEXECUTEINFO) = False Then Call Error_API(Err.LastDllError, sLocation & "\lblPage_Click", "ShellExecuteEx")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lblPage_Click")
Resume Next
End Sub
