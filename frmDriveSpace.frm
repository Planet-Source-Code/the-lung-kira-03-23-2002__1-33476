VERSION 5.00
Begin VB.Form frmDriveSpace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive Space"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmDriveSpace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFreeClustersAvailable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtTotalFreeClusters 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtFreeSectorsAvailable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtTotalFreeSectors 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtTotalSectors 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtTotalClusters 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtFreeSpaceAvailablePercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtFreeSpaceAvailable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtTotalFreeSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtTotalSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cboDrive 
      Height          =   315
      Left            =   2520
      TabIndex        =   21
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox cboRound 
      Height          =   315
      Left            =   2520
      TabIndex        =   25
      Top             =   3480
      Width           =   1815
   End
   Begin VB.ComboBox cboOutput 
      Height          =   315
      Left            =   2520
      TabIndex        =   23
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Timer tmrDriveSpace 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   2760
   End
   Begin VB.TextBox txtTotalFreeSpacePercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblTotalClusters 
      Caption         =   "Total Clusters"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblTotalSectors 
      Caption         =   "Total Sectors"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblFreeSectorsAvailable 
      Caption         =   "Free Sectors Available"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblTotalFreeSectors 
      Caption         =   "Total Free Sectors"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblFreeClustersAvailable 
      Caption         =   "Free Clusters Available"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblTotalFreeClusters 
      Caption         =   "Total Free Clusters"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblTotalSpace 
      Caption         =   "Total Space"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblTotalFreeSpace 
      Caption         =   "Total Free Space"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblFreeSpaceAvailable 
      Caption         =   "Free Space Available"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblRound 
      Caption         =   "Round"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmDriveSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmDriveSpace"


Private Sub cboDrive_Click()
On Error GoTo VB_Error

    tmrDriveSpace.Enabled = True
    tmrDriveSpace_Timer
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboDrive_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim sDrives As String
    Dim lIncrement As Long
    
    sDrives = Left$(StrReverse(ltoa_(GetLogicalDrives(), 2)) & String$(32, "0"), 32)
    
    With cboDrive
        For lIncrement = 1 To Len(sDrives)
            If Mid$(sDrives, lIncrement, 1) = "1" Then
                .AddItem Chr$(&H40 + lIncrement) & ":\"
            End If
        Next lIncrement
    End With
    
    
    With cboOutput
        .AddItem "Bytes"
        .AddItem "Kilobytes"
        .AddItem "Megabytes"
        .AddItem "Gigabytes"
        .AddItem "Terabytes"
    End With
    
    With cboRound
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
    End With
    
    
    cboOutput.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Output"), 0, 4)
    cboRound.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Round"), 0, 5)
    
    
    If Function_Exist("kernel32.dll", "GetDiskFreeSpaceExA") = False Then
        lblTotalSpace.Enabled = False
        lblTotalFreeSpace.Enabled = False
        lblFreeSpaceAvailable.Enabled = False
        lblTotalSectors.Enabled = False
        lblTotalFreeSectors.Enabled = False
        lblFreeSectorsAvailable.Enabled = False
        lblTotalClusters.Enabled = False
        lblTotalFreeClusters.Enabled = False
        lblFreeClustersAvailable.Enabled = False
        lblDrive.Enabled = False
        cboDrive.Enabled = False
        lblOutput.Enabled = False
        cboOutput.Enabled = False
        lblRound.Enabled = False
        cboRound.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    tmrDriveSpace.Enabled = False

    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Output", cboOutput.ListIndex, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Round", cboRound.ListIndex, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub tmrDriveSpace_Timer()
On Error GoTo VB_Error
    
    If cboDrive.ListIndex < 0 Then Exit Sub
    
    
    Dim dFreeBytesAvailable As Double
    Dim dTotalNumberOfBytes As Double
    Dim dTotalNumberOfFreeBytes As Double
    
    If Function_Exist("kernel32.dll", "GetDiskFreeSpaceExA") = True Then
        Dim liFreeBytesAvailable As LARGE_INTEGER
        Dim liTotalNumberOfBytes As LARGE_INTEGER
        Dim liTotalNumberOfFreeBytes As LARGE_INTEGER
        
        If GetDiskFreeSpaceEx(cboDrive.List(cboDrive.ListIndex), liFreeBytesAvailable, liTotalNumberOfBytes, liTotalNumberOfFreeBytes) = False Then
            Call Error_API(Err.LastDllError, sLocation & "\tmrDriveSpace_Timer", "GetDiskFreeSpaceEx")
            tmrDriveSpace.Enabled = False
        Else
            tmrDriveSpace.Enabled = True
        End If
        
        dFreeBytesAvailable = int32x32_int64(liFreeBytesAvailable.LowPart, liFreeBytesAvailable.HighPart)
        dTotalNumberOfBytes = int32x32_int64(liTotalNumberOfBytes.LowPart, liTotalNumberOfBytes.HighPart)
        dTotalNumberOfFreeBytes = int32x32_int64(liTotalNumberOfFreeBytes.LowPart, liTotalNumberOfFreeBytes.HighPart)
    End If
    
    
    txtTotalSpace.Text = FormatNumber$(dTotalNumberOfBytes / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
    txtTotalFreeSpace.Text = FormatNumber$(dTotalNumberOfFreeBytes / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
    txtFreeSpaceAvailable.Text = FormatNumber$(dFreeBytesAvailable / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
    
    txtTotalFreeSpacePercentage.Text = Percentage(dTotalNumberOfFreeBytes, dTotalNumberOfBytes, 0) & "%"
    txtFreeSpaceAvailablePercentage.Text = Percentage(dFreeBytesAvailable, dTotalNumberOfBytes, 0) & "%"
    
    
    Dim lSectorsPerCluster As Long
    Dim lBytesPerSector As Long
    Dim lNumberOfFreeClusters As Long
    Dim lTotalNumberOfClusters As Long
    
    If GetDiskFreeSpace(cboDrive.List(cboDrive.ListIndex), lSectorsPerCluster, lBytesPerSector, lNumberOfFreeClusters, lTotalNumberOfClusters) = False Then
        Call Error_API(Err.LastDllError, sLocation & "\tmrDriveSpace_Timer", "GetDiskFreeSpace")
        tmrDriveSpace.Enabled = False
    Else
        tmrDriveSpace.Enabled = True
    End If
    
    Dim dSectorsPerCluster As Double
    Dim dBytesPerSector As Double
    Dim dNumberOfFreeClusters As Double
    Dim dTotalNumberOfClusters As Double
    dSectorsPerCluster = int32_uint32(lSectorsPerCluster)
    dBytesPerSector = int32_uint32(lBytesPerSector)
    dNumberOfFreeClusters = int32_uint32(lNumberOfFreeClusters)
    dTotalNumberOfClusters = int32_uint32(lTotalNumberOfClusters)
     
     
    txtFreeSectorsAvailable.Text = "0"
    txtFreeClustersAvailable.Text = "0"
    txtTotalFreeSectors.Text = "0"
    txtTotalFreeClusters.Text = "0"
    txtTotalSectors.Text = "0"
    txtTotalClusters.Text = "0"
    
    If dFreeBytesAvailable > 0 Then
        If dBytesPerSector > 0 Then
            txtFreeSectorsAvailable.Text = FormatNumber$(dFreeBytesAvailable / dBytesPerSector, 0, , , True)
            
            If dSectorsPerCluster > 0 Then
                txtFreeClustersAvailable.Text = FormatNumber$((dFreeBytesAvailable / dBytesPerSector) / dSectorsPerCluster, 0, , , True)
            End If
        End If
    End If
    If dTotalNumberOfFreeBytes > 0 Then
        If dBytesPerSector > 0 Then
            txtTotalFreeSectors.Text = FormatNumber$(dTotalNumberOfFreeBytes / dBytesPerSector, 0, , , True)
            
            If dSectorsPerCluster > 0 Then
                txtTotalFreeClusters.Text = FormatNumber$((dTotalNumberOfFreeBytes / dBytesPerSector) / dSectorsPerCluster, 0, , , True)
            End If
        End If
    End If
    If dTotalNumberOfBytes > 0 Then
        If dBytesPerSector > 0 Then
            txtTotalSectors.Text = FormatNumber$(dTotalNumberOfBytes / dBytesPerSector, 0, , , True)
            
            If dSectorsPerCluster > 0 Then
                txtTotalClusters.Text = FormatNumber$((dTotalNumberOfBytes / dBytesPerSector) / dSectorsPerCluster, 0, , , True)
            End If
        End If
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\tmrDriveSpace_Timer")
Resume Next
End Sub
