Attribute VB_Name = "tlhelp32"
Option Explicit


Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Heap32First Lib "kernel32.dll" (ByRef lphe As HEAPENTRY32, ByVal th32ProcessID As Long, ByVal th32HeapID As Long) As Boolean
Public Declare Function Heap32ListFirst Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lphl As HEAPLIST32) As Boolean
Public Declare Function Heap32ListNext Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lphl As HEAPLIST32) As Boolean
Public Declare Function Heap32Next Lib "kernel32.dll" (ByRef lphe As HEAPENTRY32) As Boolean
Public Declare Function Module32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpme As MODULEENTRY32) As Boolean
Public Declare Function Module32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpme As MODULEENTRY32) As Boolean
Public Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lppe As PROCESSENTRY32) As Boolean
Public Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lppe As PROCESSENTRY32) As Boolean
Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean


Public Const HF32_DEFAULT As Long = 1
Public Const HF32_SHARED As Long = 2

Public Const LF32_FIXED As Long = &H1
Public Const LF32_FREE As Long = &H2
Public Const LF32_MOVEABLE As Long = &H4

Public Const MAX_MODULE_NAME32 As Long = 255

Public Const TH32CS_SNAPHEAPLIST As Long = &H1
Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const TH32CS_SNAPTHREAD As Long = &H4
Public Const TH32CS_SNAPMODULE As Long = &H8
Public Const TH32CS_SNAPALL As Long = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT As Long = &H80000000
    
    
Public Type HEAPENTRY32
    dwSize As Long
    hHandle As Long
    dwAddress As Long
    dwBlockSize As Long
    dwFlags As Long
    dwLockCount As Long
    dwResvd As Long
    th32ProcessID As Long
    th32HeapID As Long
End Type

Public Type HEAPLIST32
    dwSize As Long
    th32ProcessID As Long
    th32HeapID As Long
    dwFlags As Long
End Type

Public Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256    'MAX_MODULE_NAME32 + 1
    szExePath As String * MAX_PATH
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID  As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type
