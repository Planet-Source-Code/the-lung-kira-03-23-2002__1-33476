VERSION 5.00
Begin VB.Form frmWindowsInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Info"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "frmWindowsInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPersonal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   36
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkEnterprise 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   34
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkSmallBusinessRestricted 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   40
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkTerminal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   42
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkSmallBusiness 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   38
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkBackOffice 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   30
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkDataCenter 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   32
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkServer 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   27
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chkDomainController 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   25
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox chkWorkstation 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   8040
      TabIndex        =   23
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtServicePackVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtPlatformID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtCSDVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.CheckBox chkPlus 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   46
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox chkPenExtensions 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   44
      Top             =   2400
      Width           =   255
   End
   Begin VB.CheckBox chkRemoteSession 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chkSecurity 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox txtProdID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CheckBox chkDBCS 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox chkDebug 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtBoot 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblPersonal 
      Caption         =   "Personal"
      Height          =   255
      Left            =   4320
      TabIndex        =   35
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblEnterprise 
      Caption         =   "Enterprise"
      Height          =   255
      Left            =   4320
      TabIndex        =   33
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblSmallBusinessRestricted 
      Caption         =   "Small Business Restricted"
      Height          =   255
      Left            =   4320
      TabIndex        =   39
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblTerminal 
      Caption         =   "Terminal"
      Height          =   255
      Left            =   4320
      TabIndex        =   41
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblSmallBusiness 
      Caption         =   "Small Business"
      Height          =   255
      Left            =   4320
      TabIndex        =   37
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblBackOffice 
      Caption         =   "Back Office"
      Height          =   255
      Left            =   4320
      TabIndex        =   29
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblDataCenter 
      Caption         =   "Data Center"
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblSuite 
      Caption         =   "Suite"
      Height          =   255
      Left            =   4320
      TabIndex        =   28
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblProductType 
      Caption         =   "Product Type"
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblServer 
      Caption         =   "Server"
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblDomainController 
      Caption         =   "Domain Controller"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblWorkstation 
      Caption         =   "Workstation"
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblServicePackVersion 
      Caption         =   "Service Pack Version"
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblPlus 
      Caption         =   "Plus!"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblPenExtensions 
      Caption         =   "Pen Extensions"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblRemoteSession 
      Caption         =   "Remote Session"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblSecurity 
      Caption         =   "Security"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblCSDVersion 
      Caption         =   "CSD Version"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblProdID 
      Caption         =   "Product ID"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblDBCS 
      Caption         =   "DBCS Version"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblDebug 
      Caption         =   "Debug Version"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblBoot 
      Caption         =   "Boot Method"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblPlatformID 
      Caption         =   "Platform ID"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmWindowsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmWindowsSettings"


Private Sub Form_Load()
On Error GoTo VB_Error

    Dim OSVERSIONINFO As OSVERSIONINFO
    OSVERSIONINFO.dwOSVersionInfoSize = Len(OSVERSIONINFO)
    
    If GetVersionEx(OSVERSIONINFO) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "GetVersionEx")
    
    
    If WinVersion(-1, 5000000, True) = True Then
        Dim OSVERSIONINFOEX As OSVERSIONINFOEX
        OSVERSIONINFOEX.dwOSVersionInfoSize = Len(OSVERSIONINFOEX)
        
        If GetVersionEx(OSVERSIONINFOEX) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "GetVersionEx")
        
        
        With OSVERSIONINFOEX
            txtServicePackVersion.Text = .wServicePackMajor & "." & .wServicePackMinor
            
            If .wProductType And VER_NT_WORKSTATION Then chkWorkstation.value = 1
            If .wProductType And VER_NT_DOMAIN_CONTROLLER Then chkDomainController.value = 1
            If .wProductType And VER_NT_SERVER Then chkServer.value = 1
            
            If .wSuiteMask And VER_SUITE_BACKOFFICE Then chkBackOffice.value = 1
            If .wSuiteMask And VER_SUITE_DATACENTER Then chkDataCenter.value = 1
            If .wSuiteMask And VER_SUITE_ENTERPRISE Then chkEnterprise.value = 1
            If .wSuiteMask And VER_SUITE_SMALLBUSINESS Then chkSmallBusiness.value = 1
            If .wSuiteMask And VER_SUITE_SMALLBUSINESS_RESTRICTED Then chkSmallBusinessRestricted.value = 1
        End With
    Else
        lblServicePackVersion.Enabled = False
        lblProductType.Enabled = False
        lblWorkstation.Enabled = False
        lblDomainController.Enabled = False
        lblServer.Enabled = False
        lblSuite.Enabled = False
        lblBackOffice.Enabled = False
        lblDataCenter.Enabled = False
        lblEnterprise.Enabled = False
        lblSmallBusiness.Enabled = False
        lblSmallBusinessRestricted.Enabled = False
    End If
    
    
    If WinVersion(0, -1, True) = True Then
        txtName.Text = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "Version")
        txtPlatformID.Text = "WIN32 WINDOWS"
        txtProdID.Text = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductId")
        lblRemoteSession.Enabled = False
    Else
        txtName.Text = Reg_Read(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
        txtPlatformID.Text = "WIN32 NT"
        txtProdID.Text = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "ProductId")
        chkRemoteSession.value = GetSystemMetrics(SM_REMOTESESSION)
    End If
    
    
    Select Case GetSystemMetrics(SM_CLEANBOOT)
        Case 0: txtBoot.Text = "0 - Normal Boot"
        Case 1: txtBoot.Text = "1 - Fail-safe boot"
        Case 2: txtBoot.Text = "2 - Fail-safe with network boot"
        Case Else: txtBoot.Text = GetSystemMetrics(SM_CLEANBOOT) & " - Unknown"
    End Select
    

    txtCSDVersion.Text = OSVERSIONINFO.szCSDVersion
    chkDBCS.value = GetSystemMetrics(SM_DBCSENABLED)
    chkDebug.value = GetSystemMetrics(SM_DEBUG)
    chkPenExtensions.value = GetSystemMetrics(SM_PENWINDOWS)
    
    If WinVersion(4010000, -1, False) = True Then
        Dim bPlus As Boolean
        If SystemParametersInfo(SPI_GETWINDOWSEXTENSION, 0&, bPlus, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkPlus.value = bPlus
    Else
        lblPlus.Enabled = False
    End If
    
    chkSecurity.value = GetSystemMetrics(SM_SECURE)
    
    If lWinID = VER_PLATFORM_WIN32_WINDOWS Then
        txtVersion.Text = OSVERSIONINFO.dwMajorVersion & "." & OSVERSIONINFO.dwMinorVersion & "." & LOWORD(OSVERSIONINFO.dwBuildNumber)
    Else
        txtVersion.Text = OSVERSIONINFO.dwMajorVersion & "." & OSVERSIONINFO.dwMinorVersion & "." & OSVERSIONINFO.dwBuildNumber
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
