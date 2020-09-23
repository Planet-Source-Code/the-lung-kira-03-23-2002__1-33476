VERSION 5.00
Begin VB.Form frmFileSystemSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File System Settings"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmFileSystemSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkWriteBehindCache 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   25
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox chkSoftwareCompatibilityMode 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkAsyncFileCommit 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkPreserveLongNames 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox chkVirtualHDIRQ 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkForceRealModeIO 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkExtendedCharIn83Name 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   7560
      TabIndex        =   29
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtMftZoneReservation 
      Height          =   285
      Left            =   6480
      TabIndex        =   33
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox chkWin95TruncatedExtensions 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkWin31FileSystem 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   3360
      Width           =   255
   End
   Begin VB.TextBox txtS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "S"
      Top             =   1560
      Width           =   135
   End
   Begin VB.TextBox txtMaximumTunnelEntryAge 
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtMaximumTunnelEntries 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtReadAheadThreshold 
      Height          =   285
      Left            =   2520
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtContiguousFileAllocationSize 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox chkLastAccessUpdate 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   7560
      TabIndex        =   31
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chk83FileName 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   7560
      TabIndex        =   27
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   6840
      TabIndex        =   34
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblWriteBehindCache 
      Caption         =   "Write Behind Cache"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblSoftwareCompatibilityMode 
      Caption         =   "Software Compatibility Mode"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblAsyncFileCommit 
      Caption         =   "Async File Commit"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblPreserveLongNames 
      Caption         =   "Preserve Long Names"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblVirtualHDIRQ 
      Caption         =   "Virtual HD IRQ"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblForceRealModeIO 
      Caption         =   "Force Real Mode IO"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblExtendedCharIn83Name 
      Caption         =   "Extended Char In 8.3 Name"
      Height          =   255
      Left            =   4080
      TabIndex        =   28
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblMftZoneReservation 
      Caption         =   "Mft Zone Reservation"
      Height          =   255
      Left            =   4080
      TabIndex        =   32
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblContiguousFileAllocationSize 
      Caption         =   "Contiguous File Allocation Size"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblWin31FileSystem 
      Caption         =   "Win31 File System"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lblMaximumTunnelEntryAge 
      Caption         =   "Maximum Tunnel Entry Age"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblMaximumTunnelEntries 
      Caption         =   "Maximum Tunnel Entries"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblReadAheadThreshold 
      Caption         =   "Read Ahead Threshold"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblWin95TruncatedExtensions 
      Caption         =   "Win95 Truncated Extensions"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lblLastAccessUpdate 
      Caption         =   "Last Access Update"
      Height          =   255
      Left            =   4080
      TabIndex        =   30
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblNTFS 
      Caption         =   "NTFS"
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lbl83FileName 
      Caption         =   "8.3 FileName"
      Height          =   255
      Left            =   4080
      TabIndex        =   26
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmFileSystemSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmFileSystemSettings"


Private Sub chk83FileName_Click()
On Error GoTo VB_Error

    If lbl83FileName.Enabled = False Then lbl83FileName.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chk83FileName_Click")
Resume Next
End Sub

Private Sub chkAsyncFileCommit_Click()
On Error GoTo VB_Error

    If lblAsyncFileCommit.Enabled = False Then lblAsyncFileCommit.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkAsyncFileCommit_Click")
Resume Next
End Sub

Private Sub chkExtendedCharIn83Name_Click()
On Error GoTo VB_Error

    If lblExtendedCharIn83Name.Enabled = False Then lblExtendedCharIn83Name.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkExtendedCharIn83Name_Click")
Resume Next
End Sub

Private Sub chkForceRealModeIO_Click()
On Error GoTo VB_Error

    If lblForceRealModeIO.Enabled = False Then lblForceRealModeIO.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkForceRealModeIO_Click")
Resume Next
End Sub

Private Sub chkLastAccessUpdate_Click()
On Error GoTo VB_Error

    If lblLastAccessUpdate.Enabled = False Then lblLastAccessUpdate.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkLastAccessUpdate_Click")
Resume Next
End Sub

Private Sub chkPreserveLongNames_Click()
On Error GoTo VB_Error

    If lblPreserveLongNames.Enabled = False Then lblPreserveLongNames.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkPreserveLongNames_Click")
Resume Next
End Sub

Private Sub chkSoftwareCompatibilityMode_Click()
On Error GoTo VB_Error

    If lblSoftwareCompatibilityMode.Enabled = False Then lblSoftwareCompatibilityMode.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkSoftwareCompatibilityMode_Click")
Resume Next
End Sub

Private Sub chkVirtualHDIRQ_Click()
On Error GoTo VB_Error

    If lblVirtualHDIRQ.Enabled = False Then lblVirtualHDIRQ.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkVirtualHDIRQ_Click")
Resume Next
End Sub

Private Sub chkWin31FileSystem_Click()
On Error GoTo VB_Error

    If lblWin31FileSystem.Enabled = False Then lblWin31FileSystem.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkWin31FileSystem_Click")
Resume Next
End Sub

Private Sub chkWin95TruncatedExtensions_Click()
On Error GoTo VB_Error

    If lblWin95TruncatedExtensions.Enabled = False Then lblWin95TruncatedExtensions.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkWin95TruncatedExtensions_Click")
Resume Next
End Sub

Private Sub chkWriteBehindCache_Click()
On Error GoTo VB_Error

    If lblWriteBehindCache.Enabled = False Then lblWriteBehindCache.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkWriteBehindCache_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtContiguousFileAllocationSize.Text = MinMax(Val(txtContiguousFileAllocationSize.Text), 0, 4294967295#)
    txtMaximumTunnelEntries.Text = MinMax(Val(txtMaximumTunnelEntries.Text), 0, 4294967295#)
    txtMaximumTunnelEntryAge.Text = MinMax(Val(txtMaximumTunnelEntryAge.Text), 1, 30)
    txtMftZoneReservation.Text = MinMax(Val(txtMftZoneReservation.Text), 1, 4)
    txtReadAheadThreshold.Text = MinMax(Val(txtReadAheadThreshold.Text), 0, 4294967295#)
    
    
    If lblContiguousFileAllocationSize.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "ContigFileAllocSize", uint32_int32(txtContiguousFileAllocationSize.Text), REG_DWORD)
    If lblReadAheadThreshold.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "ReadAheadThreshold", uint32_int32(txtReadAheadThreshold.Text), REG_DWORD)
    If lblWin31FileSystem.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "Win31FileSystem", chkWin31FileSystem.value, REG_DWORD)
    
    
    If WinVersion(0, -1, True) = True Then
        If lblAsyncFileCommit.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "AsyncFileCommit", chkAsyncFileCommit.value, REG_DWORD)
        If lblForceRealModeIO.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "ForceRMIO", Chr$(chkForceRealModeIO.value) & String$(3, 0), REG_BINARY)
        
        If lblPreserveLongNames.Enabled = True Then
            If chkPreserveLongNames.value = 0 Then
                Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "PreserveLongNames", String$(4, 0), REG_BINARY)
            Else
                Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "PreserveLongNames", String$(4, 255), REG_BINARY)
            End If
        End If
        
        If lblSoftwareCompatibilityMode.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "SoftCompatMode", Chr$(chkSoftwareCompatibilityMode.value) & String$(3, 0), REG_BINARY)
        If lblVirtualHDIRQ.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "VirtualHDIRQ", Chr$(chkVirtualHDIRQ.value) & String$(3, 0), REG_BINARY)
        If lblWriteBehindCache.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "DriveWriteBehind", Chr$(chkWriteBehindCache.value) & String$(3, 0), REG_BINARY)
    End If
    
    If WinVersion(-1, 0, True) = True Then
        If lbl83FileName.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "NtfsDisable8dot3NameCreation", chk83FileName.value, REG_DWORD)
        If lblExtendedCharIn83Name.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "NtfsAllowExtendedCharacterIn8dot3Name", chkExtendedCharIn83Name.value, REG_DWORD)
        If lblLastAccessUpdate.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "NtfsDisableLastAccessUpdate", IIf(chkLastAccessUpdate.value, 0, 1), REG_DWORD)
        If lblMaximumTunnelEntries.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "MaximumTunnelEntries", txtMaximumTunnelEntries.Text, REG_DWORD)
        If lblMaximumTunnelEntryAge.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "MaximumTunnelEntryAgeInSeconds", txtMaximumTunnelEntryAge.Text, REG_DWORD)
        If lblMftZoneReservation.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "MftZoneReservation", txtMftZoneReservation.Text, REG_DWORD)
        If lblWin95TruncatedExtensions.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "Win95TruncatedExtensions", IIf(chkWin95TruncatedExtensions.value, 0, 1), REG_DWORD)
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim bFail As Byte
    Dim lReturn As Long
    Dim sReturn As String
    
    lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "ContigFileAllocSize", bFail)
    If bFail <> 0 Then
        lblContiguousFileAllocationSize.Enabled = False
    Else
        txtContiguousFileAllocationSize.Text = int32_uint32(lReturn)
    End If
    lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "ReadAheadThreshold", bFail)
    If bFail <> 0 Then
        lblReadAheadThreshold.Enabled = False
    Else
        txtReadAheadThreshold.Text = int32_uint32(lReturn)
    End If
    lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "Win31FileSystem", bFail)
    If bFail <> 0 Then
        lblWin31FileSystem.Enabled = False
    Else
        chkWin31FileSystem.value = IIf(lReturn, 1, 0)
    End If
    
    
    If WinVersion(0, -1, True) = True Then
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "AsyncFileCommit", bFail)
        If bFail <> 0 Then
            lblAsyncFileCommit.Enabled = False
        Else
            chkAsyncFileCommit.value = IIf(lReturn, 1, 0)
        End If
        
        sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "ForceRMIO", bFail)
        If bFail <> 0 Then
            lblForceRealModeIO.Enabled = False
        Else
            If Len(sReturn) > 0 Then chkForceRealModeIO.value = Asc(sReturn)
        End If
        
        sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "PreserveLongNames", bFail)
        If bFail <> 0 Then
            lblPreserveLongNames.Enabled = False
        Else
            If sReturn = String$(4, 255) Then chkPreserveLongNames.value = 1
        End If
        
        sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "SoftCompatMode", bFail)
        If bFail <> 0 Then
            lblSoftwareCompatibilityMode.Enabled = False
        Else
            If Len(sReturn) > 0 Then chkSoftwareCompatibilityMode.value = IIf(Asc(sReturn), 1, 0)
        End If
        
        sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "VirtualHDIRQ", bFail)
        If bFail <> 0 Then
            lblVirtualHDIRQ.Enabled = False
        Else
            If Len(sReturn) > 0 Then chkVirtualHDIRQ.value = IIf(Asc(sReturn), 1, 0)
        End If
        
        sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "DriveWriteBehind", bFail)
        If bFail <> 0 Then
            lblWriteBehindCache.Enabled = False
        Else
            If Len(sReturn) > 0 Then chkWriteBehindCache.value = IIf(Asc(sReturn), 1, 0)
        End If
    Else
        lblAsyncFileCommit.Enabled = False
        chkAsyncFileCommit.Enabled = False
        lblForceRealModeIO.Enabled = False
        chkForceRealModeIO.Enabled = False
        lblPreserveLongNames.Enabled = False
        chkPreserveLongNames.Enabled = False
        lblSoftwareCompatibilityMode.Enabled = False
        chkSoftwareCompatibilityMode.Enabled = False
        lblVirtualHDIRQ.Enabled = False
        chkVirtualHDIRQ.Enabled = False
        lblWriteBehindCache.Enabled = False
        chkWriteBehindCache.Enabled = False
    End If
    
    If WinVersion(-1, 0, True) = True Then
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "NtfsDisable8dot3NameCreation", bFail)
        If bFail <> 0 Then
            lbl83FileName.Enabled = False
        Else
            chk83FileName.value = IIf(lReturn, 1, 0)
        End If
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "NtfsAllowExtendedCharacterIn8dot3Name", bFail)
        If bFail <> 0 Then
            lblExtendedCharIn83Name.Enabled = False
        Else
            chkExtendedCharIn83Name.value = IIf(lReturn, 1, 0)
        End If
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "NtfsDisableLastAccessUpdate", bFail)
        If bFail <> 0 Then
            lblLastAccessUpdate.Enabled = False
        Else
            chkLastAccessUpdate.value = IIf(lReturn, 0, 1)
        End If
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "MaximumTunnelEntries", bFail)
        If bFail <> 0 Then
            lblMaximumTunnelEntries.Enabled = False
        Else
            txtMaximumTunnelEntries.Text = int32_uint32(lReturn)
        End If
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "MaximumTunnelEntryAgeInSeconds", bFail)
        If bFail <> 0 Then
            lblMaximumTunnelEntryAge.Enabled = False
        Else
            txtMaximumTunnelEntryAge.Text = int32_uint32(lReturn)
        End If
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "MftZoneReservation", bFail)
        If bFail <> 0 Then
            lblMftZoneReservation.Enabled = False
        Else
            txtMftZoneReservation.Text = int32_uint32(lReturn)
        End If
        
        lReturn = Reg_Read(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\FileSystem", "Win95TruncatedExtensions", bFail)
        If bFail <> 0 Then
            lblWin95TruncatedExtensions.Enabled = False
        Else
            chkWin95TruncatedExtensions.value = IIf(lReturn, 0, 1)
        End If
    Else
        lblNTFS.Enabled = False
        lbl83FileName.Enabled = False
        chk83FileName.Enabled = False
        lblExtendedCharIn83Name.Enabled = False
        chkExtendedCharIn83Name.Enabled = False
        lblLastAccessUpdate.Enabled = False
        chkLastAccessUpdate.Enabled = False
        lblMaximumTunnelEntries.Enabled = False
        txtMaximumTunnelEntries.Enabled = False
        lblMaximumTunnelEntryAge.Enabled = False
        txtMaximumTunnelEntryAge.Enabled = False
        lblMftZoneReservation.Enabled = False
        txtMftZoneReservation.Enabled = False
        lblWin95TruncatedExtensions.Enabled = False
        chkWin95TruncatedExtensions.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub txtContiguousFileAllocationSize_Change()
On Error GoTo VB_Error

    If lblContiguousFileAllocationSize.Enabled = False Then lblContiguousFileAllocationSize.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtContiguousFileAllocationSize_Change")
Resume Next
End Sub

Private Sub txtMaximumTunnelEntries_Change()
On Error GoTo VB_Error

    If lblMaximumTunnelEntries.Enabled = False Then lblMaximumTunnelEntries.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtMaximumTunnelEntries_Change")
Resume Next
End Sub

Private Sub txtMaximumTunnelEntryAge_Change()
On Error GoTo VB_Error

    If lblMaximumTunnelEntryAge.Enabled = False Then lblMaximumTunnelEntryAge.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtMaximumTunnelEntryAge_Change")
Resume Next
End Sub

Private Sub txtMftZoneReservation_Change()
On Error GoTo VB_Error

    If lblMftZoneReservation.Enabled = False Then lblMftZoneReservation.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtMftZoneReservation_Click")
Resume Next
End Sub

Private Sub txtReadAheadThreshold_Change()
On Error GoTo VB_Error

    If lblReadAheadThreshold.Enabled = False Then lblReadAheadThreshold.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtReadAheadThreshold_Change")
Resume Next
End Sub
