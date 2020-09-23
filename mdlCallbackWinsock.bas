Attribute VB_Name = "mdlCallbackWinsock"
Option Explicit


Public wsDayTime_OldProc As Long
Public wsDayTime_sockaddr As sockaddr_in
Public wsDayTime_Socket As Long

Public wsEcho_Data As String
Public wsEcho_OldProc As Long
Public wsEcho_sockaddr As sockaddr_in
Public wsEcho_Socket As Long

Public wsNameFinger_OldProc As Long
Public wsNameFinger_sockaddr As sockaddr_in
Public wsNameFinger_Socket As Long

Public wsNicnameWhois_OldProc As Long
Public wsNicnameWhois_sockaddr As sockaddr_in
Public wsNicnameWhois_Socket As Long

Public wsQOTD_OldProc As Long
Public wsQOTD_sockaddr As sockaddr_in
Public wsQOTD_Socket As Long

Public wsTime_OldProc As Long
Public wsTime_SetTime As Boolean
Public wsTime_sockaddr As sockaddr_in
Public wsTime_Socket As Long


Public wsBuffer As String
Const sLocation As String = "mdlCallbackWinsock"


Public Function wsDayTime_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error

    Select Case uMsg
        Case WM_WINSOCK_MSG
            If Forms_Loaded.bDayTime = True Then
                Select Case LOWORD(lParam)
                    Case FD_READ
                        Dim sBuffer As String
                        sBuffer = wsBuffer
                        
                        Select Case frmDayTime.cboMethod.ListIndex
                            Case 0 'UDP
                                lErrors = recvfrom(wsDayTime_Socket, sBuffer, Len(sBuffer), 0&, wsDayTime_sockaddr, Len(wsDayTime_sockaddr)): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsDayTime_Proc", "recvfrom")
                                If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                                If shutdown(wsDayTime_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsDayTime_Proc", "shutdown")
                            Case 1 'TCP
                                lErrors = recv(wsDayTime_Socket, sBuffer, Len(sBuffer), 0&): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsDayTime_Proc", "recv")
                                If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                                If shutdown(wsDayTime_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsDayTime_Proc", "shutdown")
                        End Select
                        
                        With frmDayTime
                            .txtReturned.Text = sBuffer
                            .cmdStop.Enabled = False
                            .cmdGetData.Enabled = True
                        End With
                        
                    Case FD_CLOSE
                        Call Socket_Close(wsDayTime_Socket)
                        frmDayTime.cmdStop.Enabled = False
                        frmDayTime.cmdGetData.Enabled = True
                End Select
                
                If HIWORD(lParam) <> 0 Then Call Error_API(HIWORD(lParam), sLocation & "\wsDayTime_Proc", vbNullString)
            End If
        Case Else
            wsDayTime_Proc = CallWindowProc(wsDayTime_OldProc, frmDayTime.hwnd, uMsg, wParam, lParam)
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\wsDayTime_Proc")
Resume Next
End Function

Public Function wsEcho_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error

    Select Case uMsg
        Case WM_WINSOCK_MSG
            If Forms_Loaded.bEcho = True Then
                Select Case LOWORD(lParam)
                    Case FD_READ
                        Dim sBuffer As String
                        sBuffer = wsBuffer
                        
                        Select Case frmEcho.cboMethod.ListIndex
                            Case 0 'UDP
                                lErrors = recvfrom(wsEcho_Socket, sBuffer, Len(sBuffer), 0&, wsEcho_sockaddr, Len(wsEcho_sockaddr)): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsEcho_Proc", "recvfrom")
                                If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                                If shutdown(wsEcho_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsEcho_Proc", "shutdown")
                            Case 1 'TCP
                                lErrors = recv(wsEcho_Socket, sBuffer, Len(sBuffer), 0&): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsEcho_Proc", "recv")
                                If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                                If shutdown(wsEcho_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsEcho_Proc", "shutdown")
                        End Select
                        
                        With frmEcho
                            If sBuffer = wsEcho_Data Then .chkReturnOK.value = 1
                            .cmdStop.Enabled = False
                            .cmdSendData.Enabled = True
                        End With
                        
                    Case FD_CLOSE
                        Call Socket_Close(wsEcho_Socket)
                        
                        frmEcho.cmdStop.Enabled = False
                        frmEcho.cmdSendData.Enabled = True
                        wsEcho_Data = vbNullString
                End Select
                
                If HIWORD(lParam) <> 0 Then Call Error_API(HIWORD(lParam), sLocation & "\wsEcho_Proc", vbNullString)
            End If
        Case Else
            wsEcho_Proc = CallWindowProc(wsEcho_OldProc, frmEcho.hwnd, uMsg, wParam, lParam)
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\wsEcho_Proc")
Resume Next
End Function

Public Function wsNameFinger_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error

    Select Case uMsg
        Case WM_WINSOCK_MSG
            If Forms_Loaded.bNameFinger = True Then
                Select Case LOWORD(lParam)
                    Case FD_READ
                        Dim sBuffer As String
                        sBuffer = wsBuffer
                        
                        lErrors = recv(wsNameFinger_Socket, sBuffer, Len(sBuffer), 0&): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsNameFinger_Proc", "recv")
                        If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                        frmNameFinger.txtReturned.Text = frmNameFinger.txtReturned.Text & Replace$(sBuffer, vbLf, vbCrLf, 1, -1)
                        
                    Case FD_CLOSE
                        Call Socket_Close(wsNameFinger_Socket)
                        
                        frmNameFinger.cmdStop.Enabled = False
                        frmNameFinger.cmdSendData.Enabled = True
                End Select
                
                If HIWORD(lParam) <> 0 Then Call Error_API(HIWORD(lParam), sLocation & "\wsNameFinger_Proc", vbNullString)
            End If
        Case Else
            wsNameFinger_Proc = CallWindowProc(wsNameFinger_OldProc, frmNameFinger.hwnd, uMsg, wParam, lParam)
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\wsNameFinger_Proc")
Resume Next
End Function

Public Function wsNicnameWhois_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error

    Select Case uMsg
        Case WM_WINSOCK_MSG
            If Forms_Loaded.bNicnameWhois = True Then
                Select Case LOWORD(lParam)
                    Case FD_READ
                        Dim sBuffer As String
                        sBuffer = wsBuffer
                        
                        lErrors = recv(wsNicnameWhois_Socket, sBuffer, Len(sBuffer), 0&): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsNicnameWhois_Proc", "recv")
                        If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                        frmNicnameWhois.txtReturned.Text = frmNicnameWhois.txtReturned.Text & Replace$(sBuffer, vbLf, vbCrLf, 1, -1)
                        
                    Case FD_CLOSE
                        Call Socket_Close(wsNicnameWhois_Socket)
                        
                        frmNicnameWhois.cmdStop.Enabled = False
                        frmNicnameWhois.cmdSendData.Enabled = True
                End Select
                
                If HIWORD(lParam) <> 0 Then Call Error_API(HIWORD(lParam), sLocation & "\wsNicnameWhois_Proc", vbNullString)
            End If
        Case Else
            wsNicnameWhois_Proc = CallWindowProc(wsNicnameWhois_OldProc, frmNicnameWhois.hwnd, uMsg, wParam, lParam)
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\wsNicnameWhois_Proc")
Resume Next
End Function

Public Function wsQOTD_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error

    Select Case uMsg
        Case WM_WINSOCK_MSG
            If Forms_Loaded.bQOTD = True Then
                Select Case LOWORD(lParam)
                    Case FD_READ
                        Dim sBuffer As String
                        sBuffer = wsBuffer
                        
                        Select Case frmQOTD.cboMethod.ListIndex
                            Case 0 'UDP
                                lErrors = recvfrom(wsQOTD_Socket, sBuffer, Len(sBuffer), 0&, wsQOTD_sockaddr, Len(wsQOTD_sockaddr)): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsQOTD_Proc", "recvfrom")
                                If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                                If shutdown(wsQOTD_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsQOTD_Proc", "shutdown")
                            Case 1 'TCP
                                lErrors = recv(wsQOTD_Socket, sBuffer, Len(sBuffer), 0&): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsQOTD_Proc", "recv")
                                If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                                If shutdown(wsQOTD_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsQOTD_Proc", "shutdown")
                        End Select
                        
                        With frmQOTD
                            .txtReturned.Text = sBuffer
                            .cmdStop.Enabled = False
                            .cmdGetData.Enabled = True
                        End With
                        
                    Case FD_CLOSE
                        Call Socket_Close(wsQOTD_Socket)
                        
                        frmQOTD.cmdStop.Enabled = False
                        frmQOTD.cmdGetData.Enabled = True
                End Select
                
                If HIWORD(lParam) <> 0 Then Call Error_API(HIWORD(lParam), sLocation & "\wsQOTD_Proc", vbNullString)
            End If
        Case Else
            wsQOTD_Proc = CallWindowProc(wsQOTD_OldProc, frmQOTD.hwnd, uMsg, wParam, lParam)
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\wsQOTD_Proc")
Resume Next
End Function

Public Function wsTime_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo VB_Error

    Select Case uMsg
        Case WM_WINSOCK_MSG
            If Forms_Loaded.bTime = True Then
                Select Case LOWORD(lParam)
                    Case FD_READ
                        Dim sBuffer As String
                        sBuffer = wsBuffer
                        
                        Select Case frmTime.cboMethod.ListIndex
                            Case 0 'UDP
                                lErrors = recvfrom(wsTime_Socket, sBuffer, Len(sBuffer), 0&, wsTime_sockaddr, Len(wsTime_sockaddr)): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsTime_Proc", "recvfrom")
                                If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                                If shutdown(wsTime_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsTime_Proc", "shutdown")
                            Case 1 'TCP
                                lErrors = recv(wsTime_Socket, sBuffer, Len(sBuffer), 0&): If lErrors = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsTime_Proc", "recv")
                                If lErrors > 0 Then sBuffer = Left$(sBuffer, lErrors)
                                If shutdown(wsTime_Socket, SD_BOTH) = SOCKET_ERROR Then Call Error_API(Err.LastDllError, sLocation & "\wsTime_Proc", "shutdown")
                        End Select
                        
                        
                        If Len(sBuffer) = 4 Then
                            Dim vTime As Variant
                            Dim dTime As Double
                            
                            dTime = int32_uint32(strtoul_(Right$("00" & ltoa_(Asc(Mid$(sBuffer, 1, 1)), 16), 2) & _
                                             Right$("00" & ltoa_(Asc(Mid$(sBuffer, 2, 1)), 16), 2) & _
                                             Right$("00" & ltoa_(Asc(Mid$(sBuffer, 3, 1)), 16), 2) & _
                                             Right$("00" & ltoa_(Asc(Mid$(sBuffer, 4, 1)), 16), 2), 16))
                            
                            vTime = DateAdd("s", dTime - 2208988800#, "1/1/1970")
                            If wsTime_SetTime = True Then
                                Dim SYSTEMTIME As SYSTEMTIME
                                With SYSTEMTIME
                                    .wYear = Year(vTime)
                                    .wMonth = Month(vTime)
                                    .wDay = Day(vTime)
                                    .wHour = Hour(vTime)
                                    .wMinute = Minute(vTime)
                                    .wSecond = Second(vTime)
                                    .wMilliseconds = 0
                                End With
                                
                                If SetSystemTime(SYSTEMTIME) = False Then Call Error_API(Err.LastDllError, sLocation & "\wsTime_Proc", "SetSystemTime")
                                
                                If WinVersion(0, 5000000, False) = True Then
                                    Call SendMessage(HWND_TOPMOST, WM_TIMECHANGE, 0&, 0&)
                                End If
                            End If
                            
                            
                            frmTime.txtUnFormatted.Text = dTime
                            frmTime.txtReturnedGMT.Text = vTime
                            
                            
                            Dim TIME_ZONE_INFORMATION As TIME_ZONE_INFORMATION
                            Dim lBias As Long
                            
                            Select Case GetTimeZoneInformation(TIME_ZONE_INFORMATION)
                                Case TIME_ZONE_ID_INVALID: Call Error_API(Err.LastDllError, sLocation & "\wsTime_Proc", "GetTimeZoneInformation")
                                Case TIME_ZONE_ID_UNKNOWN: Call Error_API(Err.LastDllError, sLocation & "\wsTime_Proc", "GetTimeZoneInformation")
                            End Select
                            
                            If TIME_ZONE_INFORMATION.Bias < 0 Then
                                lBias = Abs(TIME_ZONE_INFORMATION.Bias)
                            Else
                                lBias = TIME_ZONE_INFORMATION.Bias - (TIME_ZONE_INFORMATION.Bias * 2)
                            End If
                            
                            
                            vTime = DateAdd("n", lBias, vTime)
                            If frmTime.chkDaylightSavings.value = 1 Then vTime = DateAdd("n", Abs(TIME_ZONE_INFORMATION.DaylightBias), vTime)
                        End If
                        
                        With frmTime
                            .txtReturnedLocal.Text = vTime
                            .cmdSetTime.Enabled = True
                            .cmdGetData.Enabled = True
                            .cmdStop.Enabled = False
                        End With
                        
                    Case FD_CLOSE
                        Call Socket_Close(wsTime_Socket)
                        
                        frmTime.cmdStop.Enabled = False
                        frmTime.cmdGetData.Enabled = True
                        wsTime_SetTime = False
                End Select
                
                If HIWORD(lParam) <> 0 Then Call Error_API(HIWORD(lParam), sLocation & "\wsTime_Proc", vbNullString)
            End If
            
        Case Else
            wsTime_Proc = CallWindowProc(wsTime_OldProc, frmTime.hwnd, uMsg, wParam, lParam)
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\wsTime_Proc")
Resume Next
End Function
