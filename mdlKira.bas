Attribute VB_Name = "mdlKira"
Option Explicit


Public Declare Function adler32 Lib "kira_ext.dll" (ByVal adler As Long, ByVal buf As String, ByVal buf_len As Long) As Long
Public Declare Function crc32 Lib "kira_ext.dll" (ByVal crc As Long, ByVal buf As String, ByVal buf_len As Long) As Long

Public Declare Sub cpuspeed Lib "kira_ext.dll" (ByRef cpu_speed As LARGE_INTEGER)
Public Declare Sub cpuid_ Lib "kira_ext.dll" (ByVal inpEAX As Long, ByRef outEAX As Long, ByRef outEBX As Long, ByRef outECX As Long, ByRef outEDX As Long)
Public Declare Sub rdtsc Lib "kira_ext.dll" Alias "rdtsc_" (ByRef tsc As LARGE_INTEGER)

Public Declare Function ltoa Lib "kira_ext.dll" Alias "ltoa_" (ByVal value As Long, ByVal buffer As String, ByVal radix As Long) As String
Public Declare Function strtol_ Lib "kira_ext.dll" (ByVal ptr As String, ByVal radix As Long) As Long
Public Declare Function strtoul_ Lib "kira_ext.dll" (ByVal ptr As String, ByVal radix As Long) As Long
Public Declare Function ultoa Lib "kira_ext.dll" Alias "ultoa_" (ByVal value As Long, ByVal buffer As String, ByVal radix As Long) As String

Public Declare Sub Error_Message Lib "kira_ext.dll" (ByVal hObj As Long)

Public Declare Function KeyboardHook_Install Lib "kira_ext.dll" (ByVal hwnd As Long) As Long
Public Declare Sub KeyboardHook_Remove Lib "kira_ext.dll" ()
Public Declare Function MouseHook_Install Lib "kira_ext.dll" (ByVal hwnd As Long) As Long
Public Declare Sub MouseHook_Remove Lib "kira_ext.dll" ()
'Public Declare Function ShellHook_Install Lib "kira_ext.dll" (ByVal hwnd As Long) As Long
'Public Declare Sub ShellHook_Remove Lib "kira_ext.dll" ()

Const sLocation As String = "mdlKira"


Public Function CPUIDLevel_MAX() As Long
On Error GoTo VB_Error
    
    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    Call cpuid_(0, outEAX, outEBX, outECX, outEDX)
    
    CPUIDLevel_MAX = outEAX
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\CPUIDLevel_MAX")
Resume Next
End Function

Public Function CPUIDLevelExt_MAX() As Long
On Error GoTo VB_Error

    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    Call cpuid_(strtoul_("80000000", 16), outEAX, outEBX, outECX, outEDX)
    
    CPUIDLevelExt_MAX = outEAX
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\CPUIDLevelExt_MAX")
Resume Next
End Function

Public Function rdtsc_() As Double
On Error GoTo VB_Error

    Dim tsc As LARGE_INTEGER
    Call rdtsc(tsc)
    
    rdtsc_ = int32x32_int64(tsc.LowPart, tsc.HighPart)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\rdtsc_")
Resume Next
End Function

Public Function ltoa_(ByVal lValue As Long, ByVal lRadix As Long) As String
On Error GoTo VB_Error
    
    Dim sBuffer As String
    sBuffer = Space$(64)
    
    ltoa_ = CharUpper(RTrim$(Str_NullTerm_Fix(ltoa(lValue, sBuffer, lRadix))))
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\ltoa_")
Resume Next
End Function

Public Function ultoa_(ByVal lValue As Long, ByVal lRadix As Long) As String
On Error GoTo VB_Error

    Dim sBuffer As String
    sBuffer = Space$(64)
    
    ultoa_ = CharUpper(RTrim$(Str_NullTerm_Fix(ultoa(lValue, sBuffer, lRadix))))
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\ultoa_")
Resume Next
End Function


Public Sub KeyboardHookInstall()
On Error GoTo VB_Error

    If KeyboardHook = 0 Then
        Call KeyboardHook_Install(frmMain.hwnd)
    End If
    
    KeyboardHook = KeyboardHook + 1
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\KeyboardHookInstall")
Resume Next
End Sub

Public Sub MouseHookInstall()
On Error GoTo VB_Error

    If MouseHook = 0 Then
        Call MouseHook_Install(frmMain.hwnd)
    End If
    
    MouseHook = MouseHook + 1
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\MouseHookInstall")
Resume Next
End Sub

'Public Sub ShellHookInstall()
'On Error GoTo VB_Error
'
'    If ShellHook = 0 Then
'        call ShellHook_Install(frmMain.hwnd)
'    End If
'
'    ShellHook = ShellHook + 1
'
'Exit Sub
'VB_Error:
'Call Error_VB(Err, sLocation & "\ShellHookInstall")
'Resume Next
'End Sub

Public Sub KeyboardHookRemove()
On Error GoTo VB_Error

    If KeyboardHook = 1 Then
        Call KeyboardHook_Remove
    End If
    
    KeyboardHook = KeyboardHook - 1
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\KeyboardHookRemove")
Resume Next
End Sub

Public Sub MouseHookRemove()
On Error GoTo VB_Error

    If MouseHook = 1 Then
        Call MouseHook_Remove
    End If
    
    MouseHook = MouseHook - 1
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\MouseHookRemove")
Resume Next
End Sub

'Public Sub ShellHookRemove()
'On Error GoTo VB_Error
'
'    If ShellHook = 1 Then
'        Call ShellHook_Remove
'    End If
'
'    ShellHook = ShellHook - 1
'
'Exit Sub
'VB_Error:
'Call Error_VB(Err, sLocation & "\ShellHookRemove")
'Resume Next
'End Sub

Public Sub KeyboardHook_Proc(ByVal wParam As Long, ByVal lParam As Long)
On Error GoTo VB_Error
    
    Dim sLParam As String
    sLParam = StrReverse(Right$(String$(32, "0") & ltoa_(lParam, 2), 32))
    
    If Right$(sLParam, 1) = "1" Then
        KeyboardMonitor(wParam) = KeyboardMonitor(wParam) + 1
        KeyboardMonitor(256) = KeyboardMonitor(256) + 1
        
        If Forms_Loaded.bKeyboarMonitor = True Then
            frmKeyboardMonitor.lvwKeyboardMonitor.ListItems(wParam + 1).SubItems(1) = KeyboardMonitor(wParam)
            frmKeyboardMonitor.txtTotal.Text = FormatNumber(KeyboardMonitor(256), 0, , , True)
        End If
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\MouseHook_Proc")
Resume Next
End Sub

Public Sub MouseHook_Proc(ByVal uMsg As Long, ByVal lParam As Long)
On Error GoTo VB_Error
    
    If ((uMsg >= &HA0) And (uMsg <= &HAD)) Then uMsg = uMsg + 352
    
    With MouseMonitor
        Select Case uMsg
            Case WM_LBUTTONUP
                .LClicks = .LClicks + 1
                .Clicks = .Clicks + 1
                
                If Forms_Loaded.bMouseMonitor = True Then
                    frmMouseMonitor.txtLeft.Text = FormatNumber(.LClicks, 0, , , True)
                    frmMouseMonitor.txtTotalClicks.Text = FormatNumber(.Clicks, 0, , , True)
                End If
                
            Case WM_MBUTTONUP
                .MClicks = .MClicks + 1
                .Clicks = .Clicks + 1
                
                If Forms_Loaded.bMouseMonitor = True Then
                    frmMouseMonitor.txtMiddle.Text = FormatNumber(.MClicks, 0, , , True)
                    frmMouseMonitor.txtTotalClicks.Text = FormatNumber(.Clicks, 0, , , True)
                End If
                
            Case WM_RBUTTONUP
                .RClicks = .RClicks + 1
                .Clicks = .Clicks + 1
                
                If Forms_Loaded.bMouseMonitor = True Then
                    frmMouseMonitor.txtRight.Text = FormatNumber(.RClicks, 0, , , True)
                    frmMouseMonitor.txtTotalClicks.Text = FormatNumber(.Clicks, 0, , , True)
                End If
                
            Case WM_XBUTTONUP
                .XClicks = .XClicks + 1
                .Clicks = .Clicks + 1
                
                If Forms_Loaded.bMouseMonitor = True Then
                    frmMouseMonitor.txtX1.Text = FormatNumber(.XClicks, 0, , , True)
                    frmMouseMonitor.txtTotalClicks.Text = FormatNumber(.Clicks, 0, , , True)
                End If
            
            Case WM_MOUSEMOVE
                Dim lScreenEdgeX As Long
                Dim lScreenEdgeY As Long
                lScreenEdgeX = Screen.Width \ Screen.TwipsPerPixelX
                lScreenEdgeY = Screen.Height \ Screen.TwipsPerPixelY
                
                Dim POINTAPI As POINTAPI
                POINTAPI.X = LOWORD(lParam)
                POINTAPI.Y = HIWORD(lParam)
                
                If frmMain.mnuMouseMonitorOO.Checked = True Then
                    .XMovement = Abs(POINTAPI.X - .LastCoordinate.X) + .XMovement
                    .LastCoordinate.X = POINTAPI.X
                    .YMovement = Abs(POINTAPI.Y - .LastCoordinate.Y) + .YMovement
                    .LastCoordinate.Y = POINTAPI.Y
                    
                    If Forms_Loaded.bMouseMonitor = True Then
                        frmMouseMonitor.txtX.Text = FormatNumber(.XMovement, 0, , , True)
                        frmMouseMonitor.txtY.Text = FormatNumber(.YMovement, , , True)
                        frmMouseMonitor.txtTotalMovement.Text = FormatNumber(.XMovement + .YMovement, , , True)
                    End If
                End If
                
                If frmMain.mnuMouseWrapOO.Checked = True Then
                    If POINTAPI.X = lScreenEdgeX - 1 Then  'right edge reset to left
                        If SetCursorPos(1&, POINTAPI.Y) = False Then Call Error_API(Err.LastDllError, sLocation & "\frmMain_Proc", "SetCursorPos")
                        .Wrap = .Wrap + 1
                    Else
                        If POINTAPI.X = 0 Then 'left edge reset to right
                            If SetCursorPos(lScreenEdgeX - 2, POINTAPI.Y) = False Then Call Error_API(Err.LastDllError, sLocation & "\frmMain_Proc", "SetCursorPos")
                            .Wrap = .Wrap + 1
                        End If
                    End If
                    
                    If POINTAPI.Y = lScreenEdgeY - 1 Then 'bottom edge reset to top
                        If SetCursorPos(POINTAPI.X, 1&) = False Then Call Error_API(Err.LastDllError, sLocation & "\frmMain_Proc", "SetCursorPos")
                        .Wrap = .Wrap + 1
                    Else
                        If POINTAPI.Y = 0 Then 'top edge reset to bottom
                            If SetCursorPos(POINTAPI.X, lScreenEdgeY - 2) = False Then Call Error_API(Err.LastDllError, sLocation & "\frmMain_Proc", "SetCursorPos")
                            .Wrap = .Wrap + 1
                        End If
                    End If
                    
                    If Forms_Loaded.bMouseWrap = True Then
                        frmMouseWrap.txtWraps.Text = FormatNumber(.Wrap, 0, , , True)
                    End If
                End If
                
            Case WM_MOUSEWHEEL: .WheelMovement = .WheelMovement + 1
        End Select
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\MouseHook_Proc")
Resume Next
End Sub
