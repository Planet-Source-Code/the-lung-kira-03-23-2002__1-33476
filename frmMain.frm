VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kira"
   ClientHeight    =   570
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   1695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   1695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKira 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Kira"
      Height          =   350
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Begin VB.Menu mnuAccessibility 
         Caption         =   "Accessibility"
         Begin VB.Menu mnuAccessTimeout 
            Caption         =   "Access Timeout"
         End
         Begin VB.Menu mnuFilterKeys 
            Caption         =   "Filter Keys"
         End
         Begin VB.Menu mnuHighContrast 
            Caption         =   "High Contrast"
         End
         Begin VB.Menu mnuMouseKeys 
            Caption         =   "Mouse Keys"
         End
         Begin VB.Menu mnuMouseSettingsA 
            Caption         =   "Mouse Settings"
         End
         Begin VB.Menu mnuSerialKeys 
            Caption         =   "Serial Keys"
         End
         Begin VB.Menu mnuSoundSentry 
            Caption         =   "Sound Sentry"
         End
         Begin VB.Menu mnuSoundSettingsA 
            Caption         =   "Sound Settings"
         End
         Begin VB.Menu mnuStickyKeys 
            Caption         =   "Sticky Keys"
         End
         Begin VB.Menu mnuToggleKeys 
            Caption         =   "Toggle Keys"
         End
         Begin VB.Menu mnuWindowsSettingsA 
            Caption         =   "Windows Settings"
         End
      End
      Begin VB.Menu mnuCPU 
         Caption         =   "CPU"
         Begin VB.Menu mnuCPUID 
            Caption         =   "CPUID"
            Begin VB.Menu mnuCPUID00000000 
               Caption         =   "Level 0"
            End
            Begin VB.Menu mnuCPUID00000001 
               Caption         =   "Level 1"
            End
            Begin VB.Menu mnuCPUID00000002 
               Caption         =   "Level 2"
            End
            Begin VB.Menu mnuCPUID80000000 
               Caption         =   "Extended Level 0"
            End
            Begin VB.Menu mnuCPUID80000001 
               Caption         =   "Extended Level 1"
            End
            Begin VB.Menu mnuCPUID80000002_4 
               Caption         =   "Extended Level  2-4"
            End
            Begin VB.Menu mnuCPUID80000005 
               Caption         =   "Extended Level 5"
            End
            Begin VB.Menu mnuCPUID80000006 
               Caption         =   "Extended Level 6"
            End
            Begin VB.Menu mnuCPUIDOther 
               Caption         =   "Other"
            End
         End
         Begin VB.Menu mnuCPUInfo 
            Caption         =   "CPU Info"
         End
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "Drive"
         Begin VB.Menu mnuDirectories 
            Caption         =   "Directories"
         End
         Begin VB.Menu mnuFile 
            Caption         =   "File"
            Begin VB.Menu mnuFileAttributes 
               Caption         =   "File Attributes"
            End
            Begin VB.Menu mnuFileInfo 
               Caption         =   "File Info"
            End
            Begin VB.Menu mnuFileTime 
               Caption         =   "File Time"
            End
            Begin VB.Menu mnuSharedFiles 
               Caption         =   "Shared Files"
            End
         End
         Begin VB.Menu mnuFileFormat 
            Caption         =   "File Format"
            Begin VB.Menu mnuGIF 
               Caption         =   "GIF"
            End
            Begin VB.Menu mnuMZ 
               Caption         =   "MZ"
            End
            Begin VB.Menu mnuNe 
               Caption         =   "NE"
            End
            Begin VB.Menu mnuPE 
               Caption         =   "PE"
            End
         End
         Begin VB.Menu mnuFileSystemSettings 
            Caption         =   "File System Settings"
         End
         Begin VB.Menu mnuDriveInfo 
            Caption         =   "Drive Info"
         End
         Begin VB.Menu mnuDriveSpace 
            Caption         =   "Drive Space"
         End
         Begin VB.Menu mnuWindowsFileProtection 
            Caption         =   "Windows File Protection"
            Begin VB.Menu mnuWFPProtectedFiles 
               Caption         =   "Protected Files"
            End
            Begin VB.Menu mnuWFPSettings 
               Caption         =   "Settings"
            End
         End
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "Display"
         Begin VB.Menu mnuDisplayDevices 
            Caption         =   "Display Devices"
         End
         Begin VB.Menu mnuDisplayMonitors 
            Caption         =   "Display Monitors"
         End
         Begin VB.Menu mnuDisplaySettings 
            Caption         =   "Display Settings"
         End
         Begin VB.Menu mnuFont 
            Caption         =   "Font"
            Begin VB.Menu mnuFontSettings 
               Caption         =   "Font Settings"
            End
         End
         Begin VB.Menu mnuIcon 
            Caption         =   "Icon"
            Begin VB.Menu mnuIconMetrics 
               Caption         =   "Icon Metrics"
            End
            Begin VB.Menu mnuIconSettings 
               Caption         =   "Icon Settings"
            End
         End
         Begin VB.Menu mnuMenu 
            Caption         =   "Menu"
            Begin VB.Menu mnuMenuSettings 
               Caption         =   "Menu Settings"
            End
            Begin VB.Menu mnuStartMenu 
               Caption         =   "Start Menu"
            End
         End
      End
      Begin VB.Menu mnuError 
         Caption         =   "Error"
         Begin VB.Menu mnuErrorDescriptions 
            Caption         =   "Error Descriptions"
         End
         Begin VB.Menu mnuErrorSettings 
            Caption         =   "Error Settings"
         End
      End
      Begin VB.Menu mnuInternetExplorer 
         Caption         =   "Internet Explorer"
         Begin VB.Menu mnuIEHistory 
            Caption         =   "IE History"
         End
         Begin VB.Menu mnuIESettings 
            Caption         =   "IE Settings"
         End
      End
      Begin VB.Menu mnuKeyboard 
         Caption         =   "Keyboard"
         Begin VB.Menu mnuKeyboardInfo 
            Caption         =   "Keyboard Info"
         End
         Begin VB.Menu mnuKeyboardSettings 
            Caption         =   "Keyboard Settings"
         End
      End
      Begin VB.Menu mnuLocale 
         Caption         =   "Locale"
         Begin VB.Menu mnuLocalesCurrency 
            Caption         =   "Locales Currency"
         End
         Begin VB.Menu mnuLocalesDate 
            Caption         =   "Locales Date"
         End
         Begin VB.Menu mnuLocalesGeneral 
            Caption         =   "Locales General"
         End
         Begin VB.Menu mnuLocalesNumber 
            Caption         =   "Locales Number"
         End
         Begin VB.Menu mnuLocalesTime 
            Caption         =   "Locales Time"
         End
      End
      Begin VB.Menu mnuMemory 
         Caption         =   "Memory"
         Begin VB.Menu mnuMemoryInfo 
            Caption         =   "Memory Info"
         End
         Begin VB.Menu mnuMemoryStatus 
            Caption         =   "Memory Status"
         End
      End
      Begin VB.Menu mnuMouse 
         Caption         =   "Mouse"
         Begin VB.Menu mnuMouseInfo 
            Caption         =   "Mouse Info"
         End
         Begin VB.Menu mnuMouseSettings 
            Caption         =   "Mouse Settings"
         End
      End
      Begin VB.Menu mnuNetwork 
         Caption         =   "Network"
         Begin VB.Menu mnuAdaptersInfo 
            Caption         =   "Adapters Info"
         End
         Begin VB.Menu mnuCachedPasswords 
            Caption         =   "Cached Passwords"
         End
         Begin VB.Menu mnuICMP 
            Caption         =   "ICMP"
            Begin VB.Menu mnuICMPStatistics 
               Caption         =   "ICMP Statistics"
            End
         End
         Begin VB.Menu mnuIP 
            Caption         =   "IP"
            Begin VB.Menu mnuIPAddressTable 
               Caption         =   "IP Address Table"
            End
            Begin VB.Menu mnuIPForwardTable 
               Caption         =   "IP Forward Table"
            End
            Begin VB.Menu mnuIPNetTable 
               Caption         =   "IP Net Table"
            End
            Begin VB.Menu mnuIPStatisticsV4 
               Caption         =   "IP Statistics V4"
            End
            Begin VB.Menu mnuIPStatisticsV6 
               Caption         =   "IP Statistics V6"
            End
         End
         Begin VB.Menu mnuMIB2IFTable 
            Caption         =   "MIB2 IF Table"
         End
         Begin VB.Menu mnuNetworkInfo 
            Caption         =   "Network Info"
         End
         Begin VB.Menu mnuProtocolInfo 
            Caption         =   "Protocol Info"
         End
         Begin VB.Menu mnuResolveIPHost 
            Caption         =   "Resolve IP Host"
         End
         Begin VB.Menu mnuServices 
            Caption         =   "Services"
         End
         Begin VB.Menu mnuService 
            Caption         =   "Services"
            Begin VB.Menu mnuDayTime 
               Caption         =   "Day Time"
            End
            Begin VB.Menu mnuEcho 
               Caption         =   "Echo"
            End
            Begin VB.Menu mnuNameFinger 
               Caption         =   "Name Finger"
            End
            Begin VB.Menu mnuNicnameWhois 
               Caption         =   "Nicname Whois"
            End
            Begin VB.Menu mnuQOTD 
               Caption         =   "Quote of the Day"
            End
            Begin VB.Menu mnuTime 
               Caption         =   "Time"
            End
         End
         Begin VB.Menu mnuTCP 
            Caption         =   "TCP"
            Begin VB.Menu mnuTCPStatisticsV4 
               Caption         =   "TCP Statistics V4"
            End
            Begin VB.Menu mnuTCPStatisticsV6 
               Caption         =   "TCP Statistics V6"
            End
            Begin VB.Menu mnuTCPTable 
               Caption         =   "TCP Table"
            End
         End
         Begin VB.Menu mnuUDP 
            Caption         =   "UDP"
            Begin VB.Menu mnuUDPStatisticsV4 
               Caption         =   "UDP Statistics V4"
            End
            Begin VB.Menu mnuUDPStatisticsV6 
               Caption         =   "UDP Statistics V6"
            End
            Begin VB.Menu mnuUDPTable 
               Caption         =   "UDP Table"
            End
         End
      End
      Begin VB.Menu mnuPower 
         Caption         =   "Power"
         Begin VB.Menu mnuAdminPowerPolicy 
            Caption         =   "Admin Power Policy"
         End
         Begin VB.Menu mnuBatteryState 
            Caption         =   "Battery State"
         End
         Begin VB.Menu mnuPowerCapabilities 
            Caption         =   "Power Capabilities"
         End
         Begin VB.Menu mnuPowerStatus 
            Caption         =   "Power Status"
         End
         Begin VB.Menu mnuProcessorPowerInfo 
            Caption         =   "Processor Power Info"
         End
         Begin VB.Menu mnuSystemPowerInfo 
            Caption         =   "System Power Info"
         End
      End
      Begin VB.Menu mnuProcess 
         Caption         =   "Process"
         Begin VB.Menu mnuHeaps 
            Caption         =   "Heaps"
         End
         Begin VB.Menu mnuModules 
            Caption         =   "Modules"
         End
         Begin VB.Menu mnuProcessThreadTimes 
            Caption         =   "Process Thread Times"
         End
         Begin VB.Menu mnuProcesses 
            Caption         =   "Processes"
         End
         Begin VB.Menu mnuThreads 
            Caption         =   "Threads"
         End
      End
      Begin VB.Menu mnuRecycle_Bin 
         Caption         =   "Recycle Bin"
         Begin VB.Menu mnuRecycleBin 
            Caption         =   "Recycle Bin"
         End
         Begin VB.Menu mnuRecycleBinSettings 
            Caption         =   "Recycle Bin Settings"
         End
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "Window"
         Begin VB.Menu mnuWindowInfo 
            Caption         =   "Window Info"
         End
         Begin VB.Menu mnuWindowMetrics 
            Caption         =   "Window Metrics"
         End
         Begin VB.Menu mnuWindowPlacement 
            Caption         =   "Window Placement"
         End
         Begin VB.Menu mnuWindowSettings 
            Caption         =   "Window Settings"
         End
      End
      Begin VB.Menu mnuWindowS 
         Caption         =   "Windows"
         Begin VB.Menu mnuComputerName 
            Caption         =   "Computer Name"
         End
         Begin VB.Menu mnuExitWindows 
            Caption         =   "Exit Windows"
         End
         Begin VB.Menu mnuHardwareProfile 
            Caption         =   "Hardware Profile"
         End
         Begin VB.Menu mnuOperatingTime 
            Caption         =   "Operating Time"
         End
         Begin VB.Menu mnuOwner 
            Caption         =   "Owner"
         End
         Begin VB.Menu mnuStartUp 
            Caption         =   "StartUp"
         End
         Begin VB.Menu mnuWallpaper 
            Caption         =   "Wallpaper"
         End
         Begin VB.Menu mnuWindowsInfo 
            Caption         =   "Windows Info"
         End
      End
      Begin VB.Menu mnuBreak5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMonitor 
         Caption         =   "Monitor"
         Begin VB.Menu mnuKeyboardMonitor 
            Caption         =   "Keyboard Monitor"
         End
         Begin VB.Menu mnuMouseMonitor 
            Caption         =   "Mouse Monitor"
         End
         Begin VB.Menu mnuMouseWrap 
            Caption         =   "Mouse Wrap"
         End
         Begin VB.Menu mnuOnOff 
            Caption         =   "On Off"
            Begin VB.Menu mnuKeyboardMonitorOO 
               Caption         =   "Keyboard Monitor"
            End
            Begin VB.Menu mnuMouseMonitorOO 
               Caption         =   "Mouse Monitor"
            End
            Begin VB.Menu mnuMouseWrapOO 
               Caption         =   "Mouse Wrap"
            End
         End
      End
      Begin VB.Menu mnuBreak4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKira 
         Caption         =   "Kira"
         Begin VB.Menu mnuHelp 
            Caption         =   "Help"
         End
         Begin VB.Menu mnuErrorLog 
            Caption         =   "Error Log"
         End
         Begin VB.Menu mnuExtra 
            Caption         =   "Extra"
         End
         Begin VB.Menu mnuBreak3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTaskbarIcon 
            Caption         =   "Taskbar Icon"
         End
         Begin VB.Menu mnuBreak2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuResetSettings 
            Caption         =   "Reset Settings"
         End
         Begin VB.Menu mnuBreak1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCloseAll 
            Caption         =   "Close All"
         End
         Begin VB.Menu mnuMinimizeAll 
            Caption         =   "Minimize All"
         End
      End
      Begin VB.Menu mnuBreak0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMain"


Private Sub cmdKira_Click()
On Error GoTo VB_Error
    
    With frmMain
        Call .PopupMenu(.mnuMain, , .cmdKira.Left + .cmdKira.Width, cmdKira.Top)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdKira_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    Call Main
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    If bShutdown = False Then Call Main_Exit
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub mnuAccessTimeout_Click()
On Error GoTo VB_Error
    
    Call frmAccessTimeout.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuAccessTimeout_Click")
Resume Next
End Sub

Private Sub mnuAdaptersInfo_Click()
On Error GoTo VB_Error
    
    Call frmAdaptersInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuAdaptersInfo_Click")
Resume Next
End Sub

Private Sub mnuAdminPowerPolicy_Click()
On Error GoTo VB_Error
    
    Call frmAdminPowerPolicy.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuAdminPowerPolicy_Click")
Resume Next
End Sub

Private Sub mnuBatteryState_Click()
On Error GoTo VB_Error
    
    Call frmBatteryState.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuBatteryState_Click")
Resume Next
End Sub

Private Sub mnuCachedPasswords_Click()
On Error GoTo VB_Error
    
    Call frmCachedPasswords.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCachedPasswords_Click")
Resume Next
End Sub

Private Sub mnuCloseAll_Click()
On Error GoTo VB_Error

    Dim frmForm As Form
    For Each frmForm In Forms
        If Not frmForm Is frmMain Then
            Call Unload(frmForm)
        End If
    Next frmForm
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCloseAll_Click")
Resume Next
End Sub

Private Sub mnuComputerName_Click()
On Error GoTo VB_Error
    
    Call frmComputerName.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuComputerName_Click")
Resume Next
End Sub

Private Sub mnuCPUID00000000_Click()
On Error GoTo VB_Error
    
    Call frmCPUID00000000.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUID00000000_Click")
Resume Next
End Sub

Private Sub mnuCPUID00000001_Click()
On Error GoTo VB_Error
    
    Call frmCPUID00000001.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUID00000001_Click")
Resume Next
End Sub

Private Sub mnuCPUID00000002_Click()
On Error GoTo VB_Error
    
    Call frmCPUID00000002.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUID00000002_Click")
Resume Next
End Sub

Private Sub mnuCPUID80000000_Click()
On Error GoTo VB_Error
    
    Call frmCPUID80000000.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUID80000000_Click")
Resume Next
End Sub

Private Sub mnuCPUID80000001_Click()
On Error GoTo VB_Error
    
    Call frmCPUID80000001.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUID80000001_Click")
Resume Next
End Sub

Private Sub mnuCPUID80000002_4_Click()
On Error GoTo VB_Error
    
    Call frmCPUID80000002_4.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUID80000002_4_Click")
Resume Next
End Sub

Private Sub mnuCPUID80000005_Click()
On Error GoTo VB_Error
    
    Call frmCPUID80000005.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUID80000005_Click")
Resume Next
End Sub

Private Sub mnuCPUID80000006_Click()
On Error GoTo VB_Error
    
    Call frmCPUID80000006.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUID80000006_Click")
Resume Next
End Sub

Private Sub mnuCPUIDOther_Click()
On Error GoTo VB_Error
    
    Call frmCPUIDOther.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUIDOther_Click")
Resume Next
End Sub

Private Sub mnuCPUInfo_Click()
On Error GoTo VB_Error
    
    Call frmCPUInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuCPUInfo_Click")
Resume Next
End Sub

Private Sub mnuDayTime_Click()
On Error GoTo VB_Error
    
    Call frmDayTime.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuDayTime_Click")
Resume Next
End Sub

Private Sub mnuDirectories_Click()
On Error GoTo VB_Error
    
    Call frmDirectories.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuDirectories_Click")
Resume Next
End Sub

Private Sub mnuDisplayDevices_Click()
On Error GoTo VB_Error
    
    Call frmDisplayDevices.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuDisplayDevices_Click")
Resume Next
End Sub

Private Sub mnuDisplayMonitors_Click()
On Error GoTo VB_Error
    
    Call frmDisplayMonitors.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuDisplayMonitors_Click")
Resume Next
End Sub

Private Sub mnuDisplaySettings_Click()
On Error GoTo VB_Error
    
    Call frmDisplaySettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuDisplaySettings_Click")
Resume Next
End Sub

Private Sub mnuDriveInfo_Click()
On Error GoTo VB_Error
    
    Call frmDriveInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuDriveInfo_Click")
Resume Next
End Sub

Private Sub mnuDriveSpace_Click()
On Error GoTo VB_Error
    
    Call frmDriveSpace.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuDriveSpace_Click")
Resume Next
End Sub

Private Sub mnuEcho_Click()
On Error GoTo VB_Error
    
    Call frmEcho.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuEcho_Click")
Resume Next
End Sub

Private Sub mnuErrorDescriptions_Click()
On Error GoTo VB_Error
    
    Call frmErrorDescriptions.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuErrorDescriptions_Click")
Resume Next
End Sub

Private Sub mnuErrorLog_Click()
On Error GoTo VB_Error
    
    Call frmErrorLog.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuErrorLog_Click")
Resume Next
End Sub

Private Sub mnuErrorSettings_Click()
On Error GoTo VB_Error
    
    Call frmErrorSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuExtra_Click")
Resume Next
End Sub

Private Sub mnuExit_Click()
On Error GoTo VB_Error

    If bShutdown = False Then Call Main_Exit
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuExit_Click")
Resume Next
End Sub

Private Sub mnuExitWindows_Click()
On Error GoTo VB_Error
    
    Call frmExitWindows.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuExitWindows_Click")
Resume Next
End Sub

Private Sub mnuExtra_Click()
On Error GoTo VB_Error
    
    Call frmExtra.Show(, frmMain)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuExtra_Click")
Resume Next
End Sub

Private Sub mnuFileAttributes_Click()
On Error GoTo VB_Error
    
    Call frmFileAttributes.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuFileAttributes_Click")
Resume Next
End Sub

Private Sub mnuFileInfo_Click()
On Error GoTo VB_Error
    
    Call frmFileInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuFileInfo_Click")
Resume Next
End Sub

Private Sub mnuFileSystemSettings_Click()
On Error GoTo VB_Error
    
    Call frmFileSystemSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuFileSystemSettings_Click")
Resume Next
End Sub

Private Sub mnuFileTime_Click()
On Error GoTo VB_Error
    
    Call frmFileTime.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuFileTime_Click")
Resume Next
End Sub

Private Sub mnuFilterKeys_Click()
On Error GoTo VB_Error

    Call frmFilterKeys.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuFilterKeys_Click")
Resume Next
End Sub

Private Sub mnuFontSettings_Click()
On Error GoTo VB_Error

    Call frmFontSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuFontSettings_Click")
Resume Next
End Sub

Private Sub mnuGIF_Click()
On Error GoTo VB_Error
    
    Call frmGIF.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuGIF_Click")
Resume Next
End Sub

Private Sub mnuHardwareProfile_Click()
On Error GoTo VB_Error
    
    Call frmHardwareProfile.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuHardwareProfile_Click")
Resume Next
End Sub

Private Sub mnuHeaps_Click()
On Error GoTo VB_Error
    
    Call frmHeaps.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuHeaps_Click")
Resume Next
End Sub

Private Sub mnuHelp_Click()
On Error GoTo VB_Error

    Dim SHELLEXECUTEINFO As SHELLEXECUTEINFO
    With SHELLEXECUTEINFO
        .cbSize = Len(SHELLEXECUTEINFO)
        .fMask = SEE_MASK_FLAG_NO_UI
        .hwnd = frmMain.hwnd
        .lpVerb = "open"
        .lpFile = sAppPath & "\Kira.chm"
        .nShow = SW_SHOW
    End With
    
    If ShellExecuteEx(SHELLEXECUTEINFO) = False Then Call Error_API(Err.LastDllError, sLocation & "\mnuHelp_Click", "ShellExecuteEx")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuHelp_Click")
Resume Next
End Sub

Private Sub mnuHighContrast_Click()
On Error GoTo VB_Error
    
    Call frmHighContrast.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuHighContrast_Click")
Resume Next
End Sub

Private Sub mnuICMPStatistics_Click()
On Error GoTo VB_Error
    
    Call frmICMPStatistics.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuICMPStatistics_Click")
Resume Next
End Sub

Private Sub mnuIconMetrics_Click()
On Error GoTo VB_Error
    
    Call frmIconMetrics.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuIconMetrics_Click")
Resume Next
End Sub

Private Sub mnuIconSettings_Click()
On Error GoTo VB_Error
    
    Call frmIconSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuIconSettings_Click")
Resume Next
End Sub

Private Sub mnuIEHistory_Click()
On Error GoTo VB_Error
    
    Call frmIEHistory.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuIEHistory_Click")
Resume Next
End Sub

Private Sub mnuIESettings_Click()
On Error GoTo VB_Error
    
    Call frmIESettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuIESettings_Click")
Resume Next
End Sub

Private Sub mnuIPAddressTable_Click()
On Error GoTo VB_Error
    
    Call frmIPAddressTable.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuIPAddressTable_Click")
Resume Next
End Sub

Private Sub mnuIPForwardTable_Click()
On Error GoTo VB_Error
    
    Call frmIPForwardTable.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuIPForwardTable_Click")
Resume Next
End Sub

Private Sub mnuIPNetTable_Click()
On Error GoTo VB_Error
    
    Call frmIPNetTable.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuIPNetTable_Click")
Resume Next
End Sub

Private Sub mnuIPStatisticsV4_Click()
On Error GoTo VB_Error
    
    Call frmIPStatisticsV4.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuIPStatisticsV4_Click")
Resume Next
End Sub

Private Sub mnuIPStatisticsV6_Click()
On Error GoTo VB_Error
    
    Call frmIPStatisticsV6.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuIPStatisticsV6_Click")
Resume Next
End Sub

Private Sub mnuKeyboardInfo_Click()
On Error GoTo VB_Error
    
    Call frmKeyboardInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuKeyboardInfo_Click")
Resume Next
End Sub

Private Sub mnuKeyboardMonitor_Click()
On Error GoTo VB_Error
    
    Call frmKeyboardMonitor.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuKeyboardMonitor_Click")
Resume Next
End Sub

Private Sub mnuKeyboardMonitorOO_Click()
On Error GoTo VB_Error

    If mnuKeyboardMonitorOO.Checked = False Then
        mnuKeyboardMonitorOO.Checked = True
        Call KeyboardHookInstall
    Else
        mnuKeyboardMonitorOO.Checked = False
        Call KeyboardHookRemove
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuKeyboardMonitorOO_Click")
Resume Next
End Sub

Private Sub mnuKeyboardSettings_Click()
On Error GoTo VB_Error
    
    Call frmKeyboardSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuKeyboardSettings_Click")
Resume Next
End Sub

Private Sub mnuLocalesCurrency_Click()
On Error GoTo VB_Error
    
    Call frmLocalesCurrency.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuLocalesCurrency_Click")
Resume Next
End Sub

Private Sub mnuLocalesDate_Click()
On Error GoTo VB_Error
    
    Call frmLocalesDate.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuLocalesDate_Click")
Resume Next
End Sub

Private Sub mnuLocalesGeneral_Click()
On Error GoTo VB_Error
    
    Call frmLocalesGeneral.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuLocalesGeneral_Click")
Resume Next
End Sub

Private Sub mnuLocalesNumber_Click()
On Error GoTo VB_Error
    
    Call frmLocalesNumber.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuLocalesNumber_Click")
Resume Next
End Sub

Private Sub mnuLocalesTime_Click()
On Error GoTo VB_Error
    
    Call frmLocalesTime.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuLocalesTime_Click")
Resume Next
End Sub

Private Sub mnuMemoryInfo_Click()
On Error GoTo VB_Error
    
    Call frmMemoryInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMemoryInfo_Click")
Resume Next
End Sub

Private Sub mnuMemoryStatus_Click()
On Error GoTo VB_Error
    
    Call frmMemoryStatus.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMemoryStatus_Click")
Resume Next
End Sub

Private Sub mnuMenuSettings_Click()
On Error GoTo VB_Error
    
    Call frmMenuSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMenuSettings_Click")
Resume Next
End Sub

Private Sub mnuMIB2IFTable_Click()
On Error GoTo VB_Error
    
    Call frmMIB2IFTable.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMIB2IFTable_Click")
Resume Next
End Sub

Private Sub mnuMinimizeAll_Click()
On Error GoTo VB_Error

    Dim frmForm As Form
    For Each frmForm In Forms
        frmForm.WindowState = vbMinimized
    Next frmForm
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMinimizeAll_Click")
Resume Next
End Sub

Private Sub mnuModules_Click()
On Error GoTo VB_Error
    
    Call frmModules.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuModules_Click")
Resume Next
End Sub

Private Sub mnuMouseInfo_Click()
On Error GoTo VB_Error
    
    Call frmMouseInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMouseInfo_Click")
Resume Next
End Sub

Private Sub mnuMouseKeys_Click()
On Error GoTo VB_Error
    
    Call frmMouseKeys.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMouseKeys_Click")
Resume Next
End Sub

Private Sub mnuMouseMonitor_Click()
On Error GoTo VB_Error
    
    Call frmMouseMonitor.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMouseMonitor_Click")
Resume Next
End Sub

Private Sub mnuMouseMonitorOO_Click()
On Error GoTo VB_Error
    
    If mnuMouseMonitorOO.Checked = False Then
        mnuMouseMonitorOO.Checked = True
        
        Dim POINTAPI As POINTAPI
        If GetCursorPos(POINTAPI) = False Then Call Error_API(Err.LastDllError, sLocation & "\mnuMouseMonitorOO_Click", "GetCursorPos")
        
        MouseMonitor.LastCoordinate.X = POINTAPI.X
        MouseMonitor.LastCoordinate.Y = POINTAPI.Y
        
        Call MouseHookInstall
    Else
        mnuMouseMonitorOO.Checked = False
        Call MouseHookRemove
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMouseMonitorOO_Click")
Resume Next
End Sub

Private Sub mnuMouseSettings_Click()
On Error GoTo VB_Error
    
    Call frmMouseSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMouseSettings_Click")
Resume Next
End Sub

Private Sub mnuMouseSettingsA_Click()
On Error GoTo VB_Error
    
    Call frmMouseSettingsA.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMouseSettingsA_Click")
Resume Next
End Sub

Private Sub mnuMouseWrap_Click()
On Error GoTo VB_Error
    
    Call frmMouseWrap.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMouseWrap_Click")
Resume Next
End Sub

Private Sub mnuMouseWrapOO_Click()
On Error GoTo VB_Error
    
    If mnuMouseWrapOO.Checked = False Then
        mnuMouseWrapOO.Checked = True
        
        Dim POINTAPI As POINTAPI
        If GetCursorPos(POINTAPI) = False Then Call Error_API(Err.LastDllError, sLocation & "\mnuMouseWrapOO_Click", "GetCursorPos")
        
        MouseMonitor.LastCoordinate.X = POINTAPI.X
        MouseMonitor.LastCoordinate.Y = POINTAPI.Y
        
        Call MouseHookInstall
    Else
        mnuMouseWrapOO.Checked = False
        Call MouseHookRemove
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMZ_Click")
Resume Next
End Sub

Private Sub mnuMZ_Click()
On Error GoTo VB_Error
    
    Call frmMZ.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuMZ_Click")
Resume Next
End Sub

Private Sub mnuNameFinger_Click()
On Error GoTo VB_Error
    
    Call frmNameFinger.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuNameFinger_Click")
Resume Next
End Sub

Private Sub mnuNE_Click()
On Error GoTo VB_Error
    
    Call frmNE.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuNE_Click")
Resume Next
End Sub

Private Sub mnuNetworkInfo_Click()
On Error GoTo VB_Error
    
    Call frmNetworkInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuNetworkInfo_Click")
Resume Next
End Sub

Private Sub mnuNicnameWhois_Click()
On Error GoTo VB_Error
    
    Call frmNicnameWhois.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuNicnameWhois_Click")
Resume Next
End Sub

Private Sub mnuOperatingTime_Click()
On Error GoTo VB_Error
    
    Call frmOperatingTime.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuOperatingTime_Click")
Resume Next
End Sub

Private Sub mnuOwner_Click()
On Error GoTo VB_Error
    
    Call frmOwner.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuOwner_Click")
Resume Next
End Sub

Private Sub mnuPE_Click()
On Error GoTo VB_Error
    
    Call frmPE.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuPE_Click")
Resume Next
End Sub

Private Sub mnuPowerCapabilities_Click()
On Error GoTo VB_Error
    
    Call frmPowerCapabilities.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuPowerCapabilities_Click")
Resume Next
End Sub

Private Sub mnuPowerStatus_Click()
On Error GoTo VB_Error
    
    Call frmPowerStatus.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuPowerStatus_Click")
Resume Next
End Sub

Private Sub mnuProcesses_Click()
On Error GoTo VB_Error
    
    Call frmProcesses.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuProcesses_Click")
Resume Next
End Sub

Private Sub mnuProcessorPowerInfo_Click()
On Error GoTo VB_Error
    
    Call frmProcessorPowerInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuProcessorPowerInfo_Click")
Resume Next
End Sub

Private Sub mnuProcessThreadTimes_Click()
On Error GoTo VB_Error
    
    Call frmProcessThreadTimes.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuProcessThreadTimes_Click")
Resume Next
End Sub

Private Sub mnuProtocolInfo_Click()
On Error GoTo VB_Error
    
    Call frmProtocolInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuProtocolInfo_Click")
Resume Next
End Sub

Private Sub mnuQOTD_Click()
On Error GoTo VB_Error
    
    Call frmQOTD.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuQOTD_Click")
Resume Next
End Sub

Private Sub mnuRecycleBin_Click()
On Error GoTo VB_Error
    
    Call frmRecycleBin.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuRecycleBin_Click")
Resume Next
End Sub

Private Sub mnuRecycleBinSettings_Click()
On Error GoTo VB_Error
    
    Call frmRecycleBinSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuRecycleBinSettings_Click")
Resume Next
End Sub

Private Sub mnuResetSettings_Click()
On Error GoTo VB_Error

    Call Reset_Settings
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuResetSettings_Click")
Resume Next
End Sub

Private Sub mnuResolveIPHost_Click()
On Error GoTo VB_Error
    
    Call frmResolveIPHost.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuResolveIPHost_Click")
Resume Next
End Sub

Private Sub mnuSerialKeys_Click()
On Error GoTo VB_Error
    
    Call frmSerialKeys.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuSerialKeys_Click")
Resume Next
End Sub

Private Sub mnuServices_Click()
On Error GoTo VB_Error
    
    Call frmServices.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuServices_Click")
Resume Next
End Sub

Private Sub mnuSharedFiles_Click()
On Error GoTo VB_Error
    
    Call frmSharedFiles.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuSharedFiles_Click")
Resume Next
End Sub

Private Sub mnuSoundSentry_Click()
On Error GoTo VB_Error
    
    Call frmSoundSentry.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuSoundSentry_Click")
Resume Next
End Sub

Private Sub mnuSoundSettingsA_Click()
On Error GoTo VB_Error
    
    Call frmSoundSettingsA.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuSoundSettingsA_Click")
Resume Next
End Sub

Private Sub mnuStartMenu_Click()
On Error GoTo VB_Error
    
    Call frmStartMenu.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuStartMenu_Click")
Resume Next
End Sub

Private Sub mnuStartUp_Click()
On Error GoTo VB_Error
    
    Call frmStartUp.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuStartUp_Click")
Resume Next
End Sub

Private Sub mnuStickyKeys_Click()
On Error GoTo VB_Error
    
    Call frmStickyKeys.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuStickyKeys_Click")
Resume Next
End Sub

Private Sub mnuSystemPowerInfo_Click()
On Error GoTo VB_Error
    
    Call frmSystemPowerInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuSystemPowerInfo_Click")
Resume Next
End Sub

Private Sub mnuTaskbarIcon_Click()
On Error GoTo VB_Error

    If mnuTaskbarIcon.Checked = False Then
        mnuTaskbarIcon.Checked = True
        frmMain.Visible = False
        mnuMain.Visible = True
        
        Call Tray_Icon_Add
    Else
        mnuTaskbarIcon.Checked = False
        frmMain.Visible = True
        mnuMain.Visible = False
        
        Call Tray_Icon_Remove
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuTaskbarIcon_Click")
Resume Next
End Sub

Private Sub mnuTCPStatisticsV4_Click()
On Error GoTo VB_Error
    
    Call frmTCPStatisticsV4.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuTCPStatisticsV4_Click")
Resume Next
End Sub

Private Sub mnuTCPStatisticsV6_Click()
On Error GoTo VB_Error
    
    Call frmTCPStatisticsV6.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuTCPStatisticsV6_Click")
Resume Next
End Sub

Private Sub mnuTCPTable_Click()
On Error GoTo VB_Error
    
    Call frmTCPTable.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuTCPTable_Click")
Resume Next
End Sub

Private Sub mnuThreads_Click()
On Error GoTo VB_Error
    
    Call frmThreads.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuThreads_Click")
Resume Next
End Sub

Private Sub mnuTime_Click()
On Error GoTo VB_Error
    
    Call frmTime.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuTime_Click")
Resume Next
End Sub

Private Sub mnuToggleKeys_Click()
On Error GoTo VB_Error
    
    Call frmToggleKeys.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuToggleKeys_Click")
Resume Next
End Sub

Private Sub mnuUDPStatisticsV4_Click()
On Error GoTo VB_Error
    
    Call frmUDPStatisticsV4.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuUDPStatisticsV4_Click")
Resume Next
End Sub

Private Sub mnuUDPStatisticsV6_Click()
On Error GoTo VB_Error
    
    Call frmUDPStatisticsV6.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuUDPStatisticsV6_Click")
Resume Next
End Sub

Private Sub mnuUDPTable_Click()
On Error GoTo VB_Error
    
    Call frmUDPTable.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuUDPTable_Click")
Resume Next
End Sub

Private Sub mnuWallpaper_Click()
On Error GoTo VB_Error
    
    Call frmWallpaper.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuWallpaper_Click")
Resume Next
End Sub

Private Sub mnuWFPProtectedFiles_Click()
On Error GoTo VB_Error
    
    Call frmWFPProtectedFiles.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuWFPProtectedFiles_Click")
Resume Next
End Sub

Private Sub mnuWFPSettings_Click()
On Error GoTo VB_Error
    
    Call frmWFPSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuWFPSettings_Click")
Resume Next
End Sub

Private Sub mnuWindowInfo_Click()
On Error GoTo VB_Error
    
    Call frmWindowInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuWindowInfo_Click")
Resume Next
End Sub

Private Sub mnuWindowMetrics_Click()
On Error GoTo VB_Error
    
    Call frmWindowMetrics.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuWindowMetrics_Click")
Resume Next
End Sub

Private Sub mnuWindowPlacement_Click()
On Error GoTo VB_Error
    
    Call frmWindowPlacement.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuWindowPlacement_Click")
Resume Next
End Sub

Private Sub mnuWindowSettings_Click()
On Error GoTo VB_Error
    
    Call frmWindowSettings.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuWindowSettings_Click")
Resume Next
End Sub

Private Sub mnuWindowsInfo_Click()
On Error GoTo VB_Error
    
    Call frmWindowsInfo.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuWindowsInfo_Click")
Resume Next
End Sub

Private Sub mnuWindowsSettingsA_Click()
On Error GoTo VB_Error
    
    Call frmWindowsSettingsA.Show
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\mnuWindowsSettingsA_Click")
Resume Next
End Sub
