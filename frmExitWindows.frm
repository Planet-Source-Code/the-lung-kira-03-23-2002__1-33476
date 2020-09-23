VERSION 5.00
Begin VB.Form frmExitWindows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exit Windows"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmExitWindows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLock 
      Caption         =   "Lock"
      Height          =   350
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox chkForceIfHung 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkForce 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   350
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblForceIfHung 
      Caption         =   "Force If Hung"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblForce 
      Caption         =   "Force"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblMethod 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmExitWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmExitWindows"


Private Sub cmdExit_Click()
On Error GoTo VB_Error

    Dim lFlags As Long
    
    Select Case cboMethod.ListIndex
        Case 0: lFlags = EWX_LOGOFF
        Case 1: lFlags = EWX_POWEROFF
        Case 2: lFlags = EWX_REBOOT
        Case 3: lFlags = EWX_SHUTDOWN
    End Select
    
    If chkForce.value = 1 Then lFlags = lFlags Or EWX_FORCE
    If chkForceIfHung.value = 1 Then lFlags = lFlags Or EWX_FORCEIFHUNG
    
    
    If ExitWindowsEx(lFlags, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdExit_Click", "ExitWindowsEx")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdExit_Click")
Resume Next
End Sub

Private Sub cmdLock_Click()
On Error GoTo VB_Error

    If LockWorkStation() = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdLock_Click", "LockWorkStation")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdLock_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With cboMethod
        .AddItem "Logoff"
        .AddItem "Poweroff"
        .AddItem "Reboot"
        .AddItem "Shutdown"
    End With
    
    If WinVersion(-1, 5000000, True) = False Then
        lblForceIfHung.Enabled = False
        chkForceIfHung.Enabled = False
    End If
    If Function_Exist("user32.dll", "LockWorkStation") = False Then
        cmdLock.Enabled = False
    End If
    
    chkForce.value = IIf(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Force"), 1, 0)
    chkForceIfHung.value = IIf(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "ForceIfHung"), 1, 0)
    cboMethod.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Method"), 0, 3)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Force", chkForce.value, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "ForceIfHung", chkForceIfHung.value, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Method", cboMethod.ListIndex, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
