Attribute VB_Name = "mdlCallback"
Option Explicit


Const sLocation As String = "frmCallback"



Public Function frmCachedPasswords_EnumCachedPasswordsProc(PASSWORD_CACHE_ENTRY As PASSWORD_CACHE_ENTRY, ByVal lParam As Long) As Integer
On Error GoTo VB_Error

    Dim lIncrement As Integer
    Dim sResource As String
    Dim sPassword As String
    
    'PASSWORD_CACHE_ENTRY.nType
    '1 = domains
    '4 = mail/mapi clients
    '6 = RAS entries
    '19 = iexplorer entries

    For lIncrement = 1 To PASSWORD_CACHE_ENTRY.cbResource
        sResource = sResource & Chr$(PASSWORD_CACHE_ENTRY.abResource(lIncrement)) 'Combine bytes to string
    Next
    For lIncrement = PASSWORD_CACHE_ENTRY.cbResource + 1 To (PASSWORD_CACHE_ENTRY.cbResource + PASSWORD_CACHE_ENTRY.cbPassword) 'Cycle through
        sPassword = sPassword & Chr$(PASSWORD_CACHE_ENTRY.abResource(lIncrement)) 'Combine bytes to string
    Next
    
    With frmCachedPasswords.lstCachedPasswords
        .AddItem PASSWORD_CACHE_ENTRY.nType
        .AddItem sResource
        .AddItem sPassword
        .AddItem vbNullString
    End With
    
    frmCachedPasswords_EnumCachedPasswordsProc = 1 'true
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmCachedPasswords_EnumCachedPasswordsProc")
Resume Next
End Function

Public Function frmDisplayMonitor_MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByRef lprcMonitor As RECT, ByVal dwData As Long) As Boolean
On Error GoTo VB_Error
    
    frmDisplayMonitors.cboDisplayMonitors.AddItem hMonitor
    
    frmDisplayMonitor_MonitorEnumProc = 1
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmDisplayMonitor_MonitorEnumProc")
Resume Next
End Function

Public Function frmLocalesCurrency_EnumLocalesProc(ByRef lpLocaleString As Long) As Long 'Boolean
On Error GoTo VB_Error
    
    If IsBadReadPtr(lpLocaleString, 8) = True Then
        Call Error_API(Err.LastDllError, sLocation & "\frmLocalesCurrency_EnumLocalesProc", "IsBadReadPtr")
    Else
        Dim sLocale As String
        sLocale = String$(8, 0)
        Call MoveMemory(ByVal sLocale, lpLocaleString, ByVal Len(sLocale))
        
        frmLocalesCurrency.lstLocales.AddItem CharUpper(Str_NullTerm_Fix(sLocale))
    End If
    
    frmLocalesCurrency_EnumLocalesProc = 1 'True
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmLocalesCurrency_EnumLocalesProc")
Resume Next
End Function

Public Function frmLocalesDate_EnumLocalesProc(ByRef lpLocaleString As Long) As Long 'Boolean
On Error GoTo VB_Error

    If IsBadReadPtr(lpLocaleString, 8) = True Then
        Call Error_API(Err.LastDllError, sLocation & "\frmLocalesDate_EnumLocalesProc", "IsBadReadPtr")
    Else
        Dim sLocale As String
        sLocale = String$(8, 0)
        Call MoveMemory(ByVal sLocale, lpLocaleString, ByVal Len(sLocale))
        
        frmLocalesDate.lstLocales.AddItem CharUpper(Str_NullTerm_Fix(sLocale))
    End If
    
    frmLocalesDate_EnumLocalesProc = 1 'True
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmLocalesDate_EnumLocalesProc")
Resume Next
End Function

Public Function frmLocalesGeneral_EnumLocalesProc(ByRef lpLocaleString As Long) As Long 'Boolean
On Error GoTo VB_Error

    If IsBadReadPtr(lpLocaleString, 8) = True Then
        Call Error_API(Err.LastDllError, sLocation & "\frmLocalesGeneral_EnumLocalesProc", "IsBadReadPtr")
    Else
        Dim sLocale As String
        sLocale = String$(8, 0)
        Call MoveMemory(ByVal sLocale, lpLocaleString, ByVal Len(sLocale))
        
        frmLocalesGeneral.lstLocales.AddItem CharUpper(Str_NullTerm_Fix(sLocale))
    End If
    
    frmLocalesGeneral_EnumLocalesProc = 1 'True
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmLocalesGeneral_EnumLocalesProc")
Resume Next
End Function

Public Function frmLocalesNumber_EnumLocalesProc(ByRef lpLocaleString As Long) As Long 'Boolean
On Error GoTo VB_Error

    If IsBadReadPtr(lpLocaleString, 8) = True Then
        Call Error_API(Err.LastDllError, sLocation & "\frmLocalesNumber_EnumLocalesProc", "IsBadReadPtr")
    Else
        Dim sLocale As String
        sLocale = String$(8, 0)
        Call MoveMemory(ByVal sLocale, lpLocaleString, ByVal Len(sLocale))
        
        frmLocalesNumber.lstLocales.AddItem CharUpper(Str_NullTerm_Fix(sLocale))
    End If
    
    frmLocalesNumber_EnumLocalesProc = 1 'True
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmLocalesNumber_EnumLocalesProc")
Resume Next
End Function

Public Function frmLocalesTime_EnumLocalesProc(ByRef lpLocaleString As Long) As Long 'Boolean
On Error GoTo VB_Error

    If IsBadReadPtr(lpLocaleString, 8) = True Then
        Call Error_API(Err.LastDllError, sLocation & "\frmLocalesTime_EnumLocalesProc", "IsBadReadPtr")
    Else
        Dim sLocale As String
        sLocale = String$(8, 0)
        Call MoveMemory(ByVal sLocale, lpLocaleString, ByVal Len(sLocale))
        
        frmLocalesTime.lstLocales.AddItem CharUpper(Str_NullTerm_Fix(sLocale))
    End If
    
    frmLocalesTime_EnumLocalesProc = 1 'True
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmLocalesTime_EnumLocalesProc")
Resume Next
End Function

Public Function frmWindowInfo_EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error
    
    If frmWindowInfo.lvwProcess.SelectedItem Is Nothing Then Exit Function
    
    
    Dim lIncrement As Long
    Dim lProcessID As Long
    Dim lThread As Long
    
    lThread = GetWindowThreadProcessId(hwnd, lProcessID)
    
    With frmWindowInfo
        If lProcessID = .lvwProcess.SelectedItem Then
            If lThread = .lstThread.List(.lstThread.ListIndex) Then
                .lvwWindow.ListItems.Add(, , hwnd).SubItems(1) = WindowText_Get(hwnd)
            End If
        End If
    End With
    
    frmWindowInfo_EnumWindowsProc = 1
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmWindowInfo_EnumWindowsProc")
Resume Next
End Function

Public Function frmWindowPlacement_EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error
    
    If frmWindowPlacement.lvwProcess.SelectedItem Is Nothing Then Exit Function
    
    
    Dim lIncrement As Long
    Dim lProcessID As Long
    Dim lThread As Long
    
    lThread = GetWindowThreadProcessId(hwnd, lProcessID)
    
    With frmWindowPlacement
        If lProcessID = .lvwProcess.SelectedItem Then
            If lThread = .lstThread.List(.lstThread.ListIndex) Then
                .lvwWindow.ListItems.Add(, , hwnd).SubItems(1) = WindowText_Get(hwnd)
            End If
        End If
    End With
    
    frmWindowPlacement_EnumWindowsProc = 1
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmWindowPlacement_EnumWindowsProc")
Resume Next
End Function

Public Function frmWindowSettings_EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error
    
    If frmWindowSettings.lvwProcess.SelectedItem Is Nothing Then Exit Function
    
    
    Dim lIncrement As Long
    Dim lProcessID As Long
    Dim lThread As Long
    
    lThread = GetWindowThreadProcessId(hwnd, lProcessID)
    
    With frmWindowSettings
        If lProcessID = .lvwProcess.SelectedItem Then
            If lThread = .lstThread.List(.lstThread.ListIndex) Then
                .lvwWindow.ListItems.Add(, , hwnd).SubItems(1) = WindowText_Get(hwnd)
            End If
        End If
    End With
    
    frmWindowSettings_EnumWindowsProc = 1
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\frmWindowSettings_EnumWindowsProc")
Resume Next
End Function
