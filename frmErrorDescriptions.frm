VERSION 5.00
Begin VB.Form frmErrorDescriptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Descriptions"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmErrorDescriptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtErrorNumber 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get Info"
      Height          =   350
      Left            =   3840
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtDescription 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
   End
   Begin VB.ComboBox cboErrorType 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblErrorNumber 
      Caption         =   "Error Number"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblErrorType 
      Caption         =   "Error Type"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmErrorDescriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmErrorDescriptions"


Private Sub cmdGetInfo_Click()
On Error GoTo VB_Error
    
    txtErrorNumber.Text = MinMax(Val(txtErrorNumber.Text), 0, 4294967295#)
    
    
    Dim lErrorNumber As Long
    lErrorNumber = uint32_int32(txtErrorNumber.Text)
    
    Select Case cboErrorType.ListIndex
        Case 0: txtDescription.Text = API_Error(lErrorNumber)
        Case 1: txtDescription.Text = CommDlg_Error(lErrorNumber)
        Case 2: txtDescription.Text = Exception_Error(lErrorNumber)
        Case 3: txtDescription.Text = MCI_Error(lErrorNumber)
        Case 4: txtDescription.Text = PDH_Error(lErrorNumber)
    End Select
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdGetInfo_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With cboErrorType
        .AddItem "Windows"
        .AddItem "Common Dialog"
        .AddItem "General Protection Fault Exception"
        .AddItem "Media Control Interface"
        .AddItem "Performance Monitor"
    End With
    
    txtErrorNumber.Text = MinMax(Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Number")), 0, 4294967295#)
    cboErrorType.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Type"), 0, 4)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    txtErrorNumber.Text = MinMax(Val(txtErrorNumber.Text), 0, 4294967295#)
    
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Number", uint32_int32(txtErrorNumber.Text), REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Type", cboErrorType.ListIndex, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
