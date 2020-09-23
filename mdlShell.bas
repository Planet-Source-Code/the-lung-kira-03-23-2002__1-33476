Attribute VB_Name = "mdlShell"
Option Explicit


Const sLocation As String = "mdlShell"


Public Sub Tray_Icon_Add()
On Error GoTo VB_Error

    Dim NOTIFYICONDATA As NOTIFYICONDATA
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_TRAY_KIRA
        .hIcon = frmMain.Icon
        .szTip = frmMain.Caption & vbNullChar
    End With
    
    If Shell_NotifyIcon(NIM_ADD, NOTIFYICONDATA) = False Then
        If (NOTIFYICONDATA.uFlags And NIM_SETVERSION) = True Then
            Call Error_API(Err.LastDllError, sLocation & "\Tray_Icon_Add", "Shell_NotifyIcon")
        End If
    End If

Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Tray_Icon_Add")
Resume Next
End Sub

Public Sub Tray_Icon_Remove()
On Error GoTo VB_Error
    
    Dim NOTIFYICONDATA As NOTIFYICONDATA
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_TRAY_KIRA
        .hIcon = frmMain.Icon
        .szTip = vbNullChar
    End With
    
    If Shell_NotifyIcon(NIM_DELETE, NOTIFYICONDATA) = False Then
        If (NOTIFYICONDATA.uFlags And NIM_SETVERSION) = True Then
            Call Error_API(Err.LastDllError, sLocation & "\Tray_Icon_Remove", "Shell_NotifyIcon")
        End If
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Tray_Icon_Remove")
Resume Next
End Sub
