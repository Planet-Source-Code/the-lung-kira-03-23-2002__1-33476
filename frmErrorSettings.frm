VERSION 5.00
Begin VB.Form frmErrorSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Settings"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmErrorSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox cboErrorMode 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox chkWarningBeep 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblErrorMode 
      Caption         =   "Error Mode"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblWarningBeep 
      Caption         =   "Warning Beep"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmErrorSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmErrorSettings"


Private Sub cboErrorMode_Click()
On Error GoTo VB_Error
    
    If lblErrorMode.Enabled = False Then lblErrorMode.Enabled = True

Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboErrorMode_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    If WinVersion(-1, 5010000, False) = True Then
        If lblErrorMode.Enabled = True Then
            If cboErrorMode.ListIndex > -1 Then
                Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Windows", "ErrorMode", cboErrorMode.ListIndex, REG_DWORD)
            End If
        End If
    End If
    
    If SystemParametersInfo(SPI_SETBEEP, CBool(chkWarningBeep.value), 0&, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    With cboErrorMode
        .AddItem "All Error Messages"
        .AddItem "Applications Only"
        .AddItem "No Messages"
    End With
    
    If WinVersion(-1, 0, True) = True Then
        Dim bFail As Byte
        Dim lValue As Long
        
        lValue = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Windows", "ErrorMode", bFail)
        If bFail <> 0 Then
            lblErrorMode.Enabled = False
        Else
            Select Case lValue
                Case 0 To 2: cboErrorMode.ListIndex = lValue
                Case Else: cboErrorMode.ListIndex = -1
            End Select
        End If
    Else
        lblErrorMode.Enabled = False
        cboErrorMode.Enabled = False
    End If
    
    
    Dim bValue As Boolean
    If SystemParametersInfo(SPI_GETBEEP, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    chkWarningBeep.value = IIf(bValue, 1, 0)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
