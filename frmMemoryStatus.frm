VERSION 5.00
Begin VB.Form frmMemoryStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Status"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmMemoryStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAvailableExtendedVirtualMemoryPercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtAvailableExtendedVirtualMemory 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtAvailableVirtualMemory 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtAvailablePageFileMemory 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtAvailablePhysicalMemory 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Timer tmrMemoryStatus 
      Interval        =   1000
      Left            =   1440
      Top             =   2640
   End
   Begin VB.ComboBox cboRound 
      Height          =   315
      Left            =   2880
      TabIndex        =   23
      Top             =   3000
      Width           =   2295
   End
   Begin VB.ComboBox cboOutput 
      Height          =   315
      Left            =   2880
      TabIndex        =   21
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtTotalPhysicalMemory 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtAvailablePhysicalMemoryPercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtTotalPageFileMemory 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtAvailablePageFileMemoryPercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtTotalVirtualMemory 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtAvailableVirtualMemoryPercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtMemoryLoad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblRound 
      Caption         =   "Round"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblTotalPhysicalMemory 
      Caption         =   "Total Physical Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblAvailablePhysicalMemory 
      Caption         =   "Available Physical Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblTotalPageFileMemory 
      Caption         =   "Total Page File Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblAvailablePageFileMemory 
      Caption         =   "Available Page File Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblTotalVirtualMemory 
      Caption         =   "Total Virtual Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblAvailableVirtualMemory 
      Caption         =   "Available Virtual Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblAvailableExtendedVirtualMemory 
      Caption         =   "Available Extended Virtual Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblMemoryLoad 
      Caption         =   "Memory Load"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "frmMemoryStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMemoryStatus"
Dim bExtended As Boolean


Private Sub Form_Load()
On Error GoTo VB_Error

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
    
    
    bExtended = Function_Exist("kernel32.dll", "GlobalMemoryStatusEx")
    If bExtended = False Then
        lblAvailableExtendedVirtualMemory.Enabled = False
    End If
    
    
    cboOutput.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Output"), 0, 4)
    cboRound.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Round"), 0, 5)
    
    Call tmrMemoryStatus_Timer
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    tmrMemoryStatus.Enabled = False
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Output", cboOutput.ListIndex, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Round", cboRound.ListIndex, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub tmrMemoryStatus_Timer()
On Error GoTo VB_Error

    Dim dTotalPhysical As Double
    Dim dAvailablePhysical As Double
    Dim dTotalPageFile As Double
    Dim dAvailablePageFile As Double
    Dim dTotalVirtual As Double
    Dim dAvailableVirtual As Double
    Dim lMemoryLoad As Long
    
    If bExtended = True Then
        Dim MEMORYSTATUSEX As MEMORYSTATUSEX
        
        MEMORYSTATUSEX.dwLength = Len(MEMORYSTATUSEX)
        If GlobalMemoryStatusEx(MEMORYSTATUSEX) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "GlobalMemoryStatusEx")
        
        Dim dAvailableExtendedVirtual As Double
        
        With MEMORYSTATUSEX
            lMemoryLoad = .dwMemoryLoad
            dTotalPhysical = int32x32_int64(.ullTotalPhys.LowPart, .ullTotalPhys.HighPart)
            dAvailablePhysical = int32x32_int64(.ullAvailPhys.LowPart, .ullAvailPhys.HighPart)
            dTotalPageFile = int32x32_int64(.ullTotalPageFile.LowPart, .ullTotalPageFile.HighPart)
            dAvailablePageFile = int32x32_int64(.ullAvailPageFile.LowPart, .ullAvailPageFile.HighPart)
            dTotalVirtual = int32x32_int64(.ullTotalVirtual.LowPart, .ullTotalVirtual.HighPart)
            dAvailableVirtual = int32x32_int64(.ullAvailVirtual.LowPart, .ullAvailVirtual.HighPart)
            dAvailableExtendedVirtual = int32x32_int64(.ullAvailExtendedVirtual.LowPart, .ullAvailExtendedVirtual.HighPart)
        End With
        
        
        txtAvailableExtendedVirtualMemory.Text = FormatNumber$(dAvailableExtendedVirtual / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
        txtAvailableExtendedVirtualMemoryPercentage.Text = Percentage(dAvailableExtendedVirtual, dTotalVirtual, 0) & "%"
    Else
        Dim MEMORYSTATUS As MEMORYSTATUS
        Call GlobalMemoryStatus(MEMORYSTATUS)
        
        With MEMORYSTATUS
            lMemoryLoad = .dwMemoryLoad
            dTotalPhysical = .dwTotalPhys
            dAvailablePhysical = .dwAvailPhys
            dTotalPageFile = .dwTotalPageFile
            dAvailablePageFile = .dwAvailPageFile
            dTotalVirtual = .dwTotalVirtual
            dAvailableVirtual = .dwAvailVirtual
        End With
    End If
    
    
    txtTotalPhysicalMemory.Text = FormatNumber$(dTotalPhysical / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
    txtAvailablePhysicalMemory.Text = FormatNumber$(dAvailablePhysical / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
    txtTotalPageFileMemory.Text = FormatNumber$(dTotalPageFile / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
    txtAvailablePageFileMemory.Text = FormatNumber$(dAvailablePageFile / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
    txtTotalVirtualMemory.Text = FormatNumber$(dTotalVirtual / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
    txtAvailableVirtualMemory.Text = FormatNumber$(dAvailableVirtual / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
    
    
    txtAvailablePhysicalMemoryPercentage.Text = Percentage(dAvailablePhysical, dTotalPhysical, 0) & "%"
    txtAvailablePageFileMemoryPercentage.Text = Percentage(dAvailablePageFile, dTotalPageFile, 0) & "%"
    txtAvailableVirtualMemoryPercentage.Text = Percentage(dAvailableVirtual, dTotalVirtual, 0) & "%"
    
    
    txtMemoryLoad.Text = lMemoryLoad & "%"
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\tmrMemoryStatus_Timer")
Resume Next
End Sub
