Attribute VB_Name = "mdlMain"
Option Explicit


Public lWinVer As Long
Public lWinID As Long

Public bShutdown As Boolean
Public bWinsock As Boolean

Public lErrors As Long
Public sErrorLog As String
Public sErrorLogNum As Long
Public dCounterFrequency As Double
Public oldExceptionHandler As Long
Public old_frmMain_Proc As Long

Public Const WM_DLLERROR_KIRA As Long = WM_USER
Public Const WM_TRAY_KIRA As Long = WM_USER + 1
Public Const WM_HOOK_MOUSE As Long = WM_USER + 2
Public Const WM_HOOK_KEYBOARD As Long = WM_USER + 3
'Public Const WM_HOOK_JOURNAL As Long = WM_USER + 4
'Public Const WM_HOOK_SHELL As Long = WM_USER + 5
Public Const WM_WINSOCK_MSG As Long = WM_USER + 6

Public sAppPath As String
Public Const sAppVer As String = "03-23-2002"
Public Const sRegKey As String = "Software\Lung\Kira"


Public Forms_Loaded As Forms_Loaded
Public Type Forms_Loaded
    bDayTime As Boolean
    bEcho As Boolean
    bErrorLog As Boolean
    bKeyboarMonitor As Boolean
    bMouseMonitor As Boolean
    bMouseWrap As Boolean
    bNameFinger As Boolean
    bNicnameWhois As Boolean
    bQOTD As Boolean
    bTime As Boolean
End Type

Public KeyboardHook As Long
Public MouseHook As Long
'Public ShellHook As Long

Public KeyboardMonitor(0 To 256) As Double
Public MouseMonitor As MouseMonitor
Public Type MouseMonitor
    XMovement As Double
    YMovement As Double
    WheelMovement As Double
    LastCoordinate As POINTAPI
    LClicks As Double
    MClicks As Double
    RClicks As Double
    XClicks As Double
    Clicks As Double
    Wrap As Double
End Type

Const sLocation As String = "mdlMain"


Public Sub Main()
On Error GoTo VB_Error
    
    If App.PrevInstance = True Then Exit Sub
    'frmErrorLog.Show
    
    oldExceptionHandler = SetUnhandledExceptionFilter(AddressOf ExceptionHandler)
    If oldExceptionHandler = 0 Then oldExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
        
    Call SetErrorMode(SEM_FAILCRITICALERRORS Or SEM_NOGPFAULTERRORBOX Or SEM_NOOPENFILEERRORBOX)
    
    
    sAppPath = Str_BckSlhTerm_Fix(App.Path)
    
    
    Dim OSVERSIONINFO As OSVERSIONINFO
    OSVERSIONINFO.dwOSVersionInfoSize = Len(OSVERSIONINFO)
    If GetVersionEx(OSVERSIONINFO) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Main", "GetVersionEx")
    
    lWinID = OSVERSIONINFO.dwPlatformId
    If lWinID = VER_PLATFORM_WIN32_WINDOWS Then
        lWinVer = Right$("0" & OSVERSIONINFO.dwMajorVersion, 1) & _
                  Right$("00" & OSVERSIONINFO.dwMinorVersion, 2) & _
                  Right$("0000" & (LOWORD(OSVERSIONINFO.dwBuildNumber)), 4)
    Else
        lWinVer = Right$("0" & OSVERSIONINFO.dwMajorVersion, 1) & _
                  Right$("00" & OSVERSIONINFO.dwMinorVersion, 2) & _
                  Right$("0000" & (OSVERSIONINFO.dwBuildNumber), 4)
    End If
    
    
    If WinVersion(-1, 0, True) = True Then
        Call Adjust_Token_Priv(SE_INC_BASE_PRIORITY_NAME, SE_PRIVILEGE_ENABLED)
        Call Adjust_Token_Priv(SE_SHUTDOWN_NAME, SE_PRIVILEGE_ENABLED)
        Call Adjust_Token_Priv(SE_SYSTEM_PROFILE_NAME, SE_PRIVILEGE_ENABLED)
        Call Adjust_Token_Priv(SE_SYSTEMTIME_NAME, SE_PRIVILEGE_ENABLED)
        Call Adjust_Token_Priv(SE_TCB_NAME, SE_PRIVILEGE_ENABLED)
        Call Adjust_Token_Priv(SE_SECURITY_NAME, SE_PRIVILEGE_ENABLED)
    End If
    
    
    old_frmMain_Proc = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf frmMain_Proc): If old_frmMain_Proc = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Main", "SetWindowLong")
    
    
    Dim LARGE_INTEGER As LARGE_INTEGER
    If QueryPerformanceFrequency(LARGE_INTEGER) = False Then Call Error_API(Err.LastDllError, sLocation & "\Main", "QueryPerformanceFrequency")
    dCounterFrequency = int32x32_int64(LARGE_INTEGER.LowPart, LARGE_INTEGER.HighPart)
    If dCounterFrequency = 0 Then dCounterFrequency = 1
    
    
    Call Error_Message(AddressOf frmMain_Proc)
    Call Winsock_Start
    Call Verify_Settings
    
    
    Dim lKeys As Long
    For lKeys = 0 To 255
        KeyboardMonitor(lKeys) = Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\KeyboardMonitor", lKeys))
        KeyboardMonitor(256) = KeyboardMonitor(256) + KeyboardMonitor(lKeys)
    Next lKeys
    With MouseMonitor
        .XMovement = Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XMovement"))
        .YMovement = Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "YMovement"))
        .WheelMovement = Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "WheelMovement"))
        .LClicks = Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "LClicks"))
        .MClicks = Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "MClicks"))
        .RClicks = Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "RClicks"))
        .XClicks = Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XClicks"))
        .Clicks = (.LClicks + .MClicks + .RClicks + .XClicks)
        .Wrap = Val(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseWrap", "Wrap"))
    End With
    
    With frmMain
        If Reg_Read(HKEY_CURRENT_USER, sRegKey, "TaskbarIcon") = 0 Then
            .Visible = True
            .mnuMain.Visible = False
        Else
            .mnuTaskbarIcon.Checked = True
            .Visible = False
            Call Tray_Icon_Add
        End If
        Exit Sub
        If Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Monitor", "KeyboardMonitorOO") <> 0 Then
            .mnuKeyboardMonitorOO.Checked = True
            Call KeyboardHookInstall
        End If
        If Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseMonitorOO") <> 0 Then
            .mnuMouseMonitorOO.Checked = True
            Call MouseHookInstall
        End If
        If Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseWrapOO") <> 0 Then
            .mnuMouseWrapOO.Checked = True
            Call MouseHookInstall
        End If
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Main")
Resume Next
End Sub

Public Sub Main_Exit()
On Error GoTo VB_Error

    bShutdown = True
    
    
    If frmMain.mnuTaskbarIcon.Checked = True Then Call Tray_Icon_Remove
    
    
    Dim frmForm As Form
    For Each frmForm In Forms
        If Not frmForm Is frmMain Then
            Call Unload(frmForm)
        End If
    Next frmForm
    
    
    If KeyboardHook > 0 Then
        KeyboardHook = 1
        KeyboardHookRemove
    End If
    If MouseHook > 0 Then
        MouseHook = 1
        MouseHookRemove
    End If
    
    
    Call Winsock_Stop
    With frmMain
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey, "TaskbarIcon", IIf(.mnuTaskbarIcon.Checked, 1, 0), REG_DWORD)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Monitor", "KeyboardMonitorOO", IIf(.mnuKeyboardMonitorOO.Checked, 1, 0), REG_DWORD)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseMonitorOO", IIf(.mnuMouseMonitorOO.Checked, 1, 0), REG_DWORD)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseWrapOO", IIf(.mnuMouseWrapOO.Checked, 1, 0), REG_DWORD)
    End With
    Dim lKeys As Long
    For lKeys = 0 To 255
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\KeyboardMonitor", lKeys, KeyboardMonitor(lKeys), REG_SZ)
    Next lKeys
    With MouseMonitor
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XMovement", .XMovement, REG_SZ)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "YMovement", .YMovement, REG_SZ)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "WheelMovement", .WheelMovement, REG_SZ)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "LClicks", .LClicks, REG_SZ)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "MClicks", .MClicks, REG_SZ)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "RClicks", .RClicks, REG_SZ)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XClicks", .XClicks, REG_SZ)
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseWrap", "Wrap", MouseMonitor.Wrap, REG_SZ)
    End With
    
    
    If SetWindowLong(frmMain.hwnd, GWL_WNDPROC, old_frmMain_Proc) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Main_Exit", "SetWindowLong")
    
    
    Call SetErrorMode(0)
    Call SetUnhandledExceptionFilter(oldExceptionHandler)
    
    
    Call Unload(frmMain)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Main_Exit")
Resume Next
End Sub

Public Function frmMain_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error
    
    Select Case uMsg
        Case WM_DLLERROR_KIRA
            Dim sSource As String
            
            If IsBadReadPtr(lParam, lstrlen(lParam)) = True Then
                Call Error_API(Err.LastDllError, sLocation & "\frmMain_Proc", "IsBadReadPtr")
            Else
                sSource = String$(lstrlen(lstrlen(lParam)), 0)
                Call MoveMemory(ByVal sSource, ByVal lParam, Len(sSource))
            End If
            
            Call Error_API(wParam, "kira_ext.dll", sSource)
            
        Case WM_TRAY_KIRA
            Dim bPop As Boolean
            
            Select Case lParam
                Case WM_LBUTTONUP: bPop = True
                Case WM_MBUTTONUP: bPop = True
                Case WM_RBUTTONUP: bPop = True
            End Select
            
            If bPop = True Then
                If GetForegroundWindow() <> frmMain.hwnd Then
                    If SetForegroundWindow(frmMain.hwnd) = False Then Call Error_API(Err.LastDllError, sLocation & "\frmMain_Proc", "SetForegroundWindow")
                End If
                Call frmMain.PopupMenu(frmMain.mnuMain)
            End If
        
        Case WM_HOOK_MOUSE: Call MouseHook_Proc(wParam, lParam)
        Case WM_HOOK_KEYBOARD: Call KeyboardHook_Proc(wParam, lParam)
        Case Else: frmMain_Proc = CallWindowProc(old_frmMain_Proc, frmMain.hwnd, uMsg, wParam, lParam)
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmMain_Proc")
Resume Next
End Function


Sub Verify_Settings()
On Error GoTo VB_Error
    
    Dim bFail As Byte
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey, "Kira", bFail)
    If bFail <> 0 Then
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey, "Kira", sAppVer, REG_SZ)
        frmExtra.Show
    End If
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey, "TaskbarIcon", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey, "TaskbarIcon", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Monitor", "KeyboardMonitorOO", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Monitor", "KeyboardMonitorOO", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseMonitorOO", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseMonitorOO", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseWrapOO", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseWrapOO", 0, REG_DWORD)
    
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\CPUIDOther", "Level", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\CPUIDOther", "Level", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DayTime", "HostIP", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DayTime", "HostIP", vbNullString, REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Method", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Method", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Port", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Port", 13, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "GlobalChange", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "GlobalChange", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "Test", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "Test", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Output", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Output", 3, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Round", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Round", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Echo", "DataSize", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "DataSize", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Echo", "HostIP", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "HostIP", vbNullString, REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Echo", "Method", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "Method", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Echo", "Port", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "Port", 7, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Number", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Number", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Type", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Type", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Force", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Force", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "ForceIfHung", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "ForceIfHung", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Method", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Method", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Output", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Output", 2, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Round", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Round", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\IPAddressTable", "Sorted", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\IPAddressTable", "Sorted", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\IPForwardTable", "Sorted", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\IPForwardTable", "Sorted", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\IPNetTable", "Sorted", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\IPNetTable", "Sorted", 1, REG_DWORD)
    

    Dim lKeys As Long
    For lKeys = 0 To 255
        Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\KeyboardMonitor", lKeys, bFail)
        If bFail = True Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\KeyboardMonitor", lKeys, 0, REG_SZ)
    Next lKeys
    
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\LocalesCurrency", "Display", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesCurrency", "Display", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\LocalesDate", "Display", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesDate", "Display", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\LocalesGeneral", "Display", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesGeneral", "Display", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\LocalesNumber", "Display", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesNumber", "Display", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\LocalesTime", "Display", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesTime", "Display", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Output", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Output", 2, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Round", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Round", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MIB2IFTable", "Sorted", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MIB2IFTable", "Sorted", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XMovement", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XMovement", "0", REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "YMovement", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "YMovement", "0", REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "WheelMovement", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "WheelMovement", "0", REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "LClicks", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "LClicks", "0", REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "MClicks", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "MClicks", "0", REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "RClicks", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "RClicks", "0", REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XClicks", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XClicks", "0", REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\MouseWrap", "Wrap", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseWrap", "Wrap", "0", REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\NameFinger", "HostIP", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NameFinger", "HostIP", vbNullString, REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\NameFinger", "Port", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NameFinger", "Port", 79, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\NameFinger", "Send", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NameFinger", "Send", vbNullString, REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "HostIP", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "HostIP", vbNullString, REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Port", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Port", 43, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Send", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Send", vbNullString, REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\QOTD", "HostIP", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\QOTD", "HostIP", vbNullString, REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Method", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Method", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Port", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Port", 17, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Confirmation", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Confirmation", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "ProgressUI", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "ProgressUI", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Sound", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Sound", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "Host", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "Host", "localhost", REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "IP", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "IP", vbNullString, REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\TCPTable", "Sorted", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\TCPTable", "Sorted", 1, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Time", "DaylightSavings", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "DaylightSavings", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Time", "HostIP", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "HostIP", vbNullString, REG_SZ)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Time", "Method", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "Method", 0, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\Time", "Port", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "Port", 37, REG_DWORD)
    
    Call Reg_Read(HKEY_CURRENT_USER, sRegKey & "\UDPTable", "Sorted", bFail)
    If bFail <> 0 Then Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\UDPTable", "Sorted", 1, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Verify_Settings")
Resume Next
End Sub

Public Sub Reset_Settings()
On Error GoTo VB_Error
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey, "Kira", sAppVer, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey, "TaskbarIcon", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Monitor", "KeyboardMonitorOO", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseMonitorOO", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Monitor", "MouseWrapOO", 0, REG_DWORD)
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\CPUID_Other", "Level", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DayTime", "HostIP", vbNullString, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Method", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DayTime", "Port", 13, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "GlobalChange", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DisplaySettings", "Test", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Output", 3, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\DriveSpace", "Round", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "DataSize", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "HostIP", vbNullString, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "Method", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Echo", "Port", 7, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Number", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ErrorDescriptions", "Type", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Force", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "ForceIfHung", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ExitWindows", "Method", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Output", 2, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Round", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\IPAddressTable", "Sorted", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\IPForwardTable", "Sorted", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\IPNetTable", "Sorted", 1, REG_DWORD)

    Dim lKeys As Long
    For lKeys = 0 To 255
        Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\KeyboardMonitor", lKeys, 0, REG_SZ)
    Next lKeys
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesCurrency", "Display", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesDate", "Display", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesGeneral", "Display", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesNumber", "Display", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesTime", "Display", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Output", 2, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MemoryStatus", "Round", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MIB2IFTable", "Sorted", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XMovement", "0", REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "YMovement", "0", REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "WheelMovement", "0", REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "LClicks", "0", REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "MClicks", "0", REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "RClicks", "0", REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseMonitor", "XClicks", "0", REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\MouseWrap", "Wrap", "0", REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NameFinger", "HostIP", vbNullString, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NameFinger", "Port", 79, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NameFinger", "Send", vbNullString, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "HostIP", vbNullString, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Port", 43, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\NicnameWhois", "Send", vbNullString, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\QOTD", "HostIP", vbNullString, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Method", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\QOTD", "Port", 17, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Confirmation", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "ProgressUI", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\RecycleBin", "Sound", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "Host", "localhost", REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\ResolveIPHost", "IP", vbNullString, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\TCPTable", "Sorted", 1, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "DaylightSavings", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "HostIP", vbNullString, REG_SZ)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "Method", 0, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\Time", "Port", 37, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\UDPTable", "Sorted", 1, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Reset_Settings")
Resume Next
End Sub
