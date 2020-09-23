VERSION 5.00
Begin VB.Form frmRecycleBinSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recycle Bin Settings"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frmRecycleBinSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3240
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtInfoTip 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtDisplayName 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.CheckBox chkProperties 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chkDelete 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox chkRename 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chkUseRecycleBin 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblDesktopIcon 
      Caption         =   "Desktop Icon"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblInfoTip 
      Caption         =   "Info Tip"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblDisplayName 
      Caption         =   "Display Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblProperties 
      Caption         =   "Properties"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblRename 
      Caption         =   "Rename"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblUseRecycleBin 
      Caption         =   "Use Recycle Bin"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmRecycleBinSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmRecycleBinSettings"


Private Sub chkDelete_Click()
On Error GoTo VB_Error

    If lblDelete.Enabled = False Then lblDelete.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkDelete_Click")
Resume Next
End Sub

Private Sub chkProperties_Click()
On Error GoTo VB_Error

    If lblProperties.Enabled = False Then lblProperties.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkProperties_Click")
Resume Next
End Sub

Private Sub chkRename_Click()
On Error GoTo VB_Error

    If lblRename.Enabled = False Then lblRename.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkRename_Click")
Resume Next
End Sub

Private Sub chkUseRecycleBin_Click()
On Error GoTo VB_Error

    If lblUseRecycleBin.Enabled = False Then lblUseRecycleBin.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkUseRecycleBin_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error

    If lblDisplayName.Enabled = True Then Call Reg_Write(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", vbNullString, txtDisplayName.Text, REG_SZ)
    If lblInfoTip.Enabled = True Then Call Reg_Write(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", txtInfoTip.Text, REG_SZ)
    
    
    Dim lValue As Long
    Dim sInput As String
    
    If chkRename.value = 1 Then lValue = lValue + 10
    If chkDelete.value = 1 Then lValue = lValue + 20
    If chkProperties.value = 1 Then lValue = lValue + 40
    
    sInput = Reg_Read(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes")
    If Len(sInput) >= 1 Then
        sInput = Chr$(strtoul_(lValue, 16)) & Right$(sInput, Len(sInput) - 1)
        Call Reg_Write(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes", sInput, REG_BINARY)
    End If
    
    If WinVersion(4010000, 5000000, True) = True Then
        If lblUseRecycleBin.Enabled = True Then
            Call Reg_Write(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\BitBucket", "NukeOnDelete", IIf(chkUseRecycleBin.value, 0, 1), REG_DWORD)
        End If
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    Dim bFail As Byte
    Dim sReturn As String
    Dim lReturn As Long
    
    sReturn = Reg_Read(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", vbNullString, bFail)
    If bFail <> 0 Then
        lblDisplayName.Enabled = False
    Else
        txtDisplayName.Text = sReturn
    End If
    
    sReturn = Reg_Read(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", bFail)
    If bFail <> 0 Then
        lblInfoTip.Enabled = False
    Else
        txtInfoTip.Text = sReturn
    End If
    
    
    sReturn = Reg_Read(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes", bFail)
    If bFail <> 0 Then
        lblRename.Enabled = False
        lblDelete.Enabled = False
        lblProperties.Enabled = False
    Else
        If Len(sReturn) >= 4 Then
            sReturn = Right$("00" & ltoa_(Asc(Mid$(sReturn, 1, 1)), 16), 2) & _
                      Right$("00" & ltoa_(Asc(Mid$(sReturn, 2, 1)), 16), 2) & _
                      Right$("00" & ltoa_(Asc(Mid$(sReturn, 3, 1)), 16), 2) & _
                      Right$("00" & ltoa_(Asc(Mid$(sReturn, 4, 1)), 16), 2)
        
            Select Case sReturn
                Case "10010020"
                    chkRename.value = 1
                Case "20010020"
                    chkDelete.value = 1
                Case "30010020"
                    chkRename.value = 1
                    chkDelete.value = 1
                Case "40010020"
                    chkProperties.value = 1
                Case "50010020"
                    chkRename.value = 1
                    chkProperties.value = 1
                Case "60010020"
                    chkDelete.value = 1
                    chkProperties.value = 1
                Case "70010020"
                    chkRename.value = 1
                    chkDelete.value = 1
                    chkProperties.value = 1
            End Select
        End If
    End If
    
    If WinVersion(4010000, 5000000, True) = True Then
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\BitBucket", "NukeOnDelete", bFail)
        If bFail <> 0 Then
            lblUseRecycleBin.Enabled = False
        Else
            chkUseRecycleBin.value = IIf(lReturn, 0, 1)
        End If
    Else
        lblUseRecycleBin.Enabled = False
        chkUseRecycleBin.Enabled = False
    End If

Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub txtDisplayName_Change()
On Error GoTo VB_Error
    
    If lblDisplayName.Enabled = False Then lblDisplayName.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtDisplayName_Change")
Resume Next
End Sub

Private Sub txtInfoTip_Click()
On Error GoTo VB_Error

    If lblInfoTip.Enabled = False Then lblInfoTip.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtInfoTip_Click")
Resume Next
End Sub
