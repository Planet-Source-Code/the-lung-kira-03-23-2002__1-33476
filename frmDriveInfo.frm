VERSION 5.00
Begin VB.Form frmDriveInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive Info"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmDriveInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileSystemName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtMaximumComponentLength 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtVolumeSerialNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtBytesPerSector 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtSectorsPerCluster 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtVolumeName 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2400
      TabIndex        =   15
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox chkCaseIsPreserved 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkCaseSensitive 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkUnicodeStoredonDisk 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkPersistantACLS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   24
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkFileCompression 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   26
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkVolumeisCompressed 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   28
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkNamedStreams 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   30
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkReadOnlyVolume 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   32
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkSupportsEncryption 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   34
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkSupportsObjectIDs 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   36
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkSupportsReparsePoints 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   38
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkSupportsSparseFiles 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   40
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkVolumeQuotas 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   42
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3000
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtDriveType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox cboDrive 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblVolumeName 
      Caption         =   "Volume Name"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblVolumeSerialNumber 
      Caption         =   "Volume Serial Number"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblFileSystemName 
      Caption         =   "File System Name"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblMaximumComponentLength 
      Caption         =   "Maximum Component Length"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblCaseIsPreserved 
      Caption         =   "Case Is Preserved"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblCaseSensitive 
      Caption         =   "Case Sensitive"
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblUnicodeStoredonDisk 
      Caption         =   "Unicode Stored on Disk"
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblPersistantACLS 
      Caption         =   "Persistant ACLS"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblFileCompression 
      Caption         =   "File Compression"
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblVolumeisCompressed 
      Caption         =   "Volume is Compressed"
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblNamedStreams 
      Caption         =   "Named Streams"
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblReadOnlyVolume 
      Caption         =   "Read Only Volume"
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblSupportsEncryption 
      Caption         =   "Supports Encryption"
      Height          =   255
      Left            =   4200
      TabIndex        =   33
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblSupportsObjectIDs 
      Caption         =   "Supports Object IDs"
      Height          =   255
      Left            =   4200
      TabIndex        =   35
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblSupportsReparsePoints 
      Caption         =   "Supports Reparse Points"
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblSupportsSparseFiles 
      Caption         =   "Supports Sparse Files"
      Height          =   255
      Left            =   4200
      TabIndex        =   39
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblVolumeQuotas 
      Caption         =   "Volume Quotas"
      Height          =   255
      Left            =   4200
      TabIndex        =   41
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblSectorsPerCluster 
      Caption         =   "Sectors Per Cluster"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblBytesPerSector 
      Caption         =   "Bytes Per Sector"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblDriveType 
      Caption         =   "Drive Type"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmDriveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmDriveInfo"


Private Sub cboDrive_Click()
On Error GoTo VB_Error
    
    Dim lDriveType As Long
    lDriveType = GetDriveType(cboDrive.List(cboDrive.ListIndex))
    
    Select Case lDriveType
        Case DRIVE_UNKNOWN: txtDriveType.Text = "Unknown"
        Case DRIVE_NO_ROOT_DIR: txtDriveType.Text = "No Root Directory"
        Case DRIVE_REMOVABLE: txtDriveType.Text = "Removable"
        Case DRIVE_FIXED: txtDriveType.Text = "Fixed"
        Case DRIVE_REMOTE: txtDriveType.Text = "Remote"
        Case DRIVE_CDROM: txtDriveType.Text = "CDROM"
        Case DRIVE_RAMDISK: txtDriveType.Text = "RAM Disk"
        Case Else: txtDriveType.Text = "Unknown " & lDriveType
    End Select
    
    
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long
    Dim lNumberOfFreeClusters As Long
    Dim lTotalNumberOfClusters As Long
    
    If GetDiskFreeSpace(cboDrive.List(cboDrive.ListIndex), lSectorsPerCluster, lBytesPerSector, lNumberOfFreeClusters, lTotalNumberOfClusters) = False Then Call Error_API(Err.LastDllError, sLocation & "\cboDrive_Click", "GetDiskFreeSpace")
    
    txtSectorsPerCluster.Text = FormatNumber$(lSectorsPerCluster, 0, , , True)
    txtBytesPerSector.Text = FormatNumber$(lBytesPerSector, 0, , , True)
    
    
    Dim sVolumeName As String
    Dim lVolumeSerialNumber As Long
    Dim lMaximumComponentLength As Long
    Dim lFileSystemFlags As Long
    Dim sFileSystemName As String
    
    sVolumeName = String$(256, 0)
    sFileSystemName = String$(256, 0)
    
    If GetVolumeInformation(cboDrive.List(cboDrive.ListIndex), sVolumeName, Len(sVolumeName), lVolumeSerialNumber, lMaximumComponentLength, lFileSystemFlags, sFileSystemName, Len(sFileSystemName)) = False Then Call Error_API(Err.LastDllError, sLocation & "\cboDrive_Click", "GetVolumeInformation")
    
    txtVolumeName.Text = sVolumeName
    txtVolumeSerialNumber.Text = Right$("00000000" & ltoa_(lVolumeSerialNumber, 16), 8)
    txtMaximumComponentLength.Text = FormatNumber$(lMaximumComponentLength, 0, , , True)
    txtFileSystemName.Text = sFileSystemName
    
    chkCaseIsPreserved.value = IIf(lFileSystemFlags And FS_CASE_IS_PRESERVED, 1, 0)
    chkCaseSensitive.value = IIf(lFileSystemFlags And FS_CASE_SENSITIVE, 1, 0)
    chkUnicodeStoredonDisk.value = IIf(lFileSystemFlags And FS_UNICODE_STORED_ON_DISK, 1, 0)
    chkPersistantACLS.value = IIf(lFileSystemFlags And FS_PERSISTENT_ACLS, 1, 0)
    chkFileCompression.value = IIf(lFileSystemFlags And FS_FILE_COMPRESSION, 1, 0)
    chkVolumeisCompressed.value = IIf(lFileSystemFlags And FS_VOL_IS_COMPRESSED, 1, 0)
    chkNamedStreams.value = IIf(lFileSystemFlags And FILE_NAMED_STREAMS, 1, 0)
    chkSupportsEncryption.value = IIf(lFileSystemFlags And FILE_SUPPORTS_ENCRYPTION, 1, 0)
    chkSupportsObjectIDs.value = IIf(lFileSystemFlags And FILE_SUPPORTS_OBJECT_IDS, 1, 0)
    chkSupportsReparsePoints.value = IIf(lFileSystemFlags And FILE_SUPPORTS_REPARSE_POINTS, 1, 0)
    chkSupportsSparseFiles.value = IIf(lFileSystemFlags And FILE_SUPPORTS_SPARSE_FILES, 1, 0)
    chkVolumeQuotas.value = IIf(lFileSystemFlags And FILE_VOLUME_QUOTAS, 1, 0)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboDrive_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    If SetVolumeLabel(cboDrive.List(cboDrive.ListIndex), txtVolumeName.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetVolumeLabel")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
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
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
