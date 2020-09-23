VERSION 5.00
Begin VB.Form frmRecycleBin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recycle Bin"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frmRecycleBin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEmpty 
      Caption         =   "Empty"
      Height          =   350
      Left            =   3240
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cboDrive 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox chkSound 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkProgressUI 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkConfirmation 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdEmptyAll 
      Caption         =   "Empty All"
      Height          =   350
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblSound 
      Caption         =   "Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblProgressUI 
      Caption         =   "Progress UI"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblConfirmation 
      Caption         =   "Confirmation"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmRecycleBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmRecycleBin"


Private Sub cmdEmpty_Click()
On Error GoTo VB_Error

    If cboDrive.ListIndex = -1 Then Exit Sub
    
    Dim lFlags As Long
    
    If chkConfirmation.value = 0 Then lFlags = lFlags Or SHERB_NOCONFIRMATION
    If chkProgressUI.value = 0 Then lFlags = lFlags Or SHERB_NOPROGRESSUI
    If chkSound.value = 0 Then lFlags = lFlags Or SHERB_NOSOUND
    
    If SHEmptyRecycleBin(0&, cboDrive.List(cboDrive.ListIndex), lFlags) <> S_OK Then Call Error_API(Err.LastDllError, sLocation & "\cmdEmpty_Click", "SHEmptyRecycleBin")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdEmpty_Click")
Resume Next
End Sub

Private Sub cmdEmptyAll_Click()
On Error GoTo VB_Error

    Dim lFlags As Long
    
    If chkConfirmation.value = 0 Then lFlags = lFlags Or SHERB_NOCONFIRMATION
    If chkProgressUI.value = 0 Then lFlags = lFlags Or SHERB_NOPROGRESSUI
    If chkSound.value = 0 Then lFlags = lFlags Or SHERB_NOSOUND
    
    If SHEmptyRecycleBin(0&, vbNullString, lFlags) <> S_OK Then Call Error_API(Err.LastDllError, sLocation & "\cmdEmptyAll_Click", "SHEmptyRecycleBin")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdEmptyAll_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim sDrives As String
    Dim lIncrement As Long
    
    sDrives = Left$(StrReverse(ltoa_(GetLogicalDrives, 2)) & String$(32, "0"), 32)
    
    With cboDrive
        For lIncrement = 1 To Len(sDrives)
            If Mid$(sDrives, lIncrement, 1) = "1" Then
                .AddItem Chr$(&H40 + lIncrement) & ":\"
            End If
        Next lIncrement
    End With
    
    
    chkConfirmation.value = IIf(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Confirmation"), 1, 0)
    chkProgressUI.value = IIf(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "ProgressUI"), 1, 0)
    chkSound.value = IIf(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Sound"), 1, 0)
    
    
    If Function_Exist("shell32.dll", "SHEmptyRecycleBinA") = False Then
        cmdEmpty.Enabled = False
        cmdEmptyAll.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Confirmation", chkConfirmation.value, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "ProgressUI", chkProgressUI.value, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Sound", chkSound.value, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
