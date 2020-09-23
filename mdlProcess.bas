Attribute VB_Name = "mdlProcess"
Option Explicit


Const sLocation As String = "mdlProcess"


Public Sub Adjust_Token_Priv(ByVal lPriv As String, ByVal lAttrib As Long)
On Error GoTo VB_Error
    
    Dim hTokenHandle As Long
    Dim tLuid As LUID
    Dim tkpNewState As TOKEN_PRIVILEGES
    Dim tkpPreviousState As TOKEN_PRIVILEGES
    Dim lBufferLength As Long
    
    If OpenProcessToken(GetCurrentProcess(), (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hTokenHandle) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Adjust_Token_Priv", "OpenProcessToken")
    If LookupPrivilegeValue(ComputerName_Get, lPriv, tLuid) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Adjust_Token_Priv", "LookupPrivilegeValue")
    
    With tkpNewState
        .PrivilegeCount = 1
        .Privileges(0).Attributes = lAttrib
        .Privileges(0).pLuid = tLuid
    End With
    
    If AdjustTokenPrivileges(hTokenHandle, False, tkpNewState, Len(tkpPreviousState), tkpPreviousState, lBufferLength) = False Then Call Error_API(Err.LastDllError, sLocation & "\Adjust_Token_Priv", "LookupPrivilegeValue")
    If CloseHandle(hTokenHandle) = False Then Call Error_API(Err.LastDllError, sLocation & "\Adjust_Token_Priv", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Adjust_Token_Priv")
Resume Next
End Sub

Public Function Heap32_Enum(ByRef Heap() As HEAPENTRY32, ByVal lProcessID As Long, ByVal lHeapID As Long) As Long
On Error GoTo VB_Error

    ReDim Heap(0)
    
    Dim HEAPENTRY32 As HEAPENTRY32
    Dim lHeap As Long
    
    
    HEAPENTRY32.dwSize = Len(HEAPENTRY32)
    If Heap32First(HEAPENTRY32, lProcessID, lHeapID) = False Then
        Heap32_Enum = -1
        Call Error_API(Err.LastDllError, sLocation & "\Heap32_Enum", "Heap32First")
        
        Exit Function
    Else
        ReDim Heap(lHeap)
        Heap(lHeap) = HEAPENTRY32
    End If
    
    Do
        If Heap32Next(HEAPENTRY32) = False Then
            Exit Do
        Else
            lHeap = lHeap + 1
            ReDim Preserve Heap(lHeap)
            Heap(lHeap) = HEAPENTRY32
        End If
        
        If bShutdown = True Then Exit Do
    Loop
    
    
    Heap32_Enum = lHeap
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Heap32_Enum")
Resume Next
End Function

Public Function Heap32List_Enum(ByRef HeapList() As HEAPLIST32, ByVal lProcessID As Long) As Long
On Error GoTo VB_Error

    ReDim HeapList(0)
    
    Dim HEAPLIST32 As HEAPLIST32
    Dim hSnapShot As Long
    Dim lHeapList As Long
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPHEAPLIST, lProcessID): If hSnapShot = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\Heap32List_Enum", "CreateToolhelp32Snapshot")
    
    HEAPLIST32.dwSize = Len(HEAPLIST32)
    If Heap32ListFirst(hSnapShot, HEAPLIST32) = False Then
        Heap32List_Enum = -1
        Call Error_API(Err.LastDllError, sLocation & "\Heap32List_Enum", "Heap32ListFirst")
        
        If CloseHandle(hSnapShot) = False Then Call Error_API(Err.LastDllError, sLocation & "\Heap32List_Enum", "CloseHandle")
        Exit Function
    Else
        ReDim HeapList(lHeapList)
        HeapList(lHeapList) = HEAPLIST32
    End If
    
    Do
        If Heap32ListNext(hSnapShot, HEAPLIST32) = False Then
            Exit Do
        Else
            lHeapList = lHeapList + 1
            ReDim Preserve HeapList(lHeapList)
            HeapList(lHeapList) = HEAPLIST32
            
            DoEvents
        End If
        
        If bShutdown = True Then Exit Do
    Loop
    
    If CloseHandle(hSnapShot) = False Then Call Error_API(Err.LastDllError, sLocation & "\Heap32List_Enum", "CloseHandle")
    
    Heap32List_Enum = lHeapList
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Heap32List_Enum")
Resume Next
End Function

Public Function Module32_Enum(ByRef Module() As MODULEENTRY32, Optional ByVal lProcessID As Long) As Long
On Error GoTo VB_Error

    ReDim Module(0)
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = True Then
        Dim MODULEENTRY32 As MODULEENTRY32
        Dim hSnapShot As Long
        Dim lModule As Long
        
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, lProcessID): If hSnapShot = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\Module32_Enum", "CreateToolhelp32Snapshot")
        
        MODULEENTRY32.dwSize = Len(MODULEENTRY32)
        If Module32First(hSnapShot, MODULEENTRY32) = False Then
            Module32_Enum = -1
            Call Error_API(Err.LastDllError, sLocation & "\Module32_Enum", "Module32First")
            
            If CloseHandle(hSnapShot) = False Then Call Error_API(Err.LastDllError, sLocation & "\Module32_Enum", "CloseHandle")
            Exit Function
        Else
            ReDim Module(lModule)
            Module(lModule) = MODULEENTRY32
        End If
        
        Do
            If Module32Next(hSnapShot, MODULEENTRY32) = False Then
                Exit Do
            Else
                lModule = lModule + 1
                ReDim Preserve Module(lModule)
                Module(lModule) = MODULEENTRY32
            End If
            
            If bShutdown = True Then Exit Do
        Loop
        
        If CloseHandle(hSnapShot) = False Then Call Error_API(Err.LastDllError, sLocation & "\Module32_Enum", "CloseHandle")
        
        Module32_Enum = lModule
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Module32_Enum")
Resume Next
End Function

Public Function Process32_Enum(ByRef Process() As PROCESSENTRY32) As Long
On Error GoTo VB_Error

    ReDim Process(0)
    
    Dim PROCESSENTRY32 As PROCESSENTRY32
    Dim hSnapShot As Long
    Dim lProcess As Long
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&): If hSnapShot = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\Process32_Enum", "CreateToolhelp32Snapshot")
    
    PROCESSENTRY32.dwSize = Len(PROCESSENTRY32)
    If Process32First(hSnapShot, PROCESSENTRY32) = False Then
        Process32_Enum = -1
        Call Error_API(Err.LastDllError, sLocation & "\Process32_Enum", "Process32First")
        
        If CloseHandle(hSnapShot) = False Then Call Error_API(Err.LastDllError, sLocation & "\Process32_Enum", "CloseHandle")
        Exit Function
    Else
        ReDim Process(lProcess)
        Process(lProcess) = PROCESSENTRY32
    End If

    Do
        If Process32Next(hSnapShot, PROCESSENTRY32) = False Then
            Exit Do
        Else
            lProcess = lProcess + 1
            ReDim Preserve Process(lProcess)
            Process(lProcess) = PROCESSENTRY32
        End If
        
        If bShutdown = True Then Exit Do
    Loop
    
    If CloseHandle(hSnapShot) = False Then Call Error_API(Err.LastDllError, sLocation & "\Process32_Enum", "CloseHandle")
    
    Process32_Enum = lProcess
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Process32_Enum")
Resume Next
End Function

Public Function Thread32_Enum(ByRef Thread() As THREADENTRY32, ByVal lProcessID As Long) As Long
On Error GoTo VB_Error

    ReDim Thread(0)
    
    Dim THREADENTRY32 As THREADENTRY32
    Dim hSnapShot As Long
    Dim lThread As Long
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lProcessID): If hSnapShot = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\Thread32_Enum", "CreateToolhelp32Snapshot")
    
    THREADENTRY32.dwSize = Len(THREADENTRY32)
    If Thread32First(hSnapShot, THREADENTRY32) = False Then
        Thread32_Enum = -1
        Call Error_API(Err.LastDllError, sLocation & "\Thread32_Enum", "Thread32First")
        
        If CloseHandle(hSnapShot) = False Then Call Error_API(Err.LastDllError, sLocation & "\Thread32_Enum", "CloseHandle")
        Exit Function
    Else
        ReDim Thread(lThread)
        Thread(lThread) = THREADENTRY32
    End If
    
    Do
        If Thread32Next(hSnapShot, THREADENTRY32) = False Then
            Exit Do
        Else
            lThread = lThread + 1
            ReDim Preserve Thread(lThread)
            Thread(lThread) = THREADENTRY32
        End If
        
        If bShutdown = True Then Exit Do
    Loop
    
    If CloseHandle(hSnapShot) = False Then Call Error_API(Err.LastDllError, sLocation & "\Thread32_Enum", "CloseHandle")
    
    Thread32_Enum = lThread
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Thread32_Enum")
Resume Next
End Function
