Attribute VB_Name = "winbase"
Option Explicit


Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Boolean, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Boolean
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean
Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function FileEncryptionStatus Lib "advapi32.dll" Alias "FileEncryptionStatusA" (ByVal lpFileName As String, ByRef lpStatus As Long) As Boolean
Public Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Boolean
Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Boolean
Public Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Boolean
Public Declare Function GetComputerNameEx Lib "kernel32.dll" Alias "GetComputerNameExA" (ByVal NameType As COMPUTER_NAME_FORMAT, ByVal lpBuffer As String, ByRef lpnSize As Long) As Boolean
Public Declare Function GetCurrentHwProfile Lib "advapi32.dll" Alias "GetCurrentHwProfileA" (ByRef lpHwProfileInfo As HW_PROFILE_INFO) As Boolean
Public Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32.dll" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, ByRef lpSectorsPerCluster As Long, ByRef lpBytesPerSector As Long, ByRef lpNumberOfFreeClusters As Long, ByRef lpTotalNumberOfClusters As Long) As Boolean
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, ByRef lpFreeBytesAvailableToCaller As LARGE_INTEGER, ByRef lpTotalNumberOfBytes As LARGE_INTEGER, ByRef lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Boolean
Public Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Boolean
Public Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Boolean
Public Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileInformationByHandle Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Boolean
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpFileSizeHigh As Long) As Long
Public Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSize As LARGE_INTEGER) As Boolean
Public Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Boolean
Public Declare Function GetFileType Lib "kernel32.dll" (ByVal hFile As Long) As Long
Public Declare Function GetLogicalDrives Lib "kernel32.dll" () As Long
Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetProcessAffinityMask Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpProcessAffinityMask As Long, ByRef lpSystemAffinityMask As Long) As Boolean
Public Declare Function GetProcessIoCounters Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpIoCounters As IO_COUNTERS) As Boolean
Public Declare Function GetProcessTimes Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpCreationTime As FILETIME, ByRef lpExitTime As FILETIME, ByRef lpKernelTime As FILETIME, ByRef lpUserTime As FILETIME) As Boolean
Public Declare Function GetProcessVersion Lib "kernel32.dll" (ByVal ProcessId As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetSystemPowerStatus Lib "kernel32.dll" (ByRef lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Boolean
Public Declare Function GetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function GetThreadTimes Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpCreationTime As FILETIME, ByRef lpExitTime As FILETIME, ByRef lpKernelTime As FILETIME, ByRef lpUserTime As FILETIME) As Boolean
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetTimeZoneInformation Lib "kernel32.dll" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Boolean
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As Any) As Boolean
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Boolean
Public Declare Function IsBadReadPtr Lib "kernel32.dll" (ByRef lp As Any, ByVal ucb As Long) As Boolean
Public Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Boolean
Public Declare Function lstrlen Lib "kernel32.dll" (ByVal lpString As Any) As Long
Public Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Boolean
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As LARGE_INTEGER) As Boolean
Public Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As LARGE_INTEGER) As Boolean
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Boolean
Public Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function SetComputerName Lib "kernel32.dll" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Boolean
Public Declare Function SetComputerNameEx Lib "kernel32.dll" Alias "SetComputerNameExA" (ByVal NameType As COMPUTER_NAME_FORMAT, ByVal lpBuffer As String) As Boolean
Public Declare Function SetErrorMode Lib "kernel32.dll" (ByVal uMode As Long) As Long
Public Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Boolean
Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Boolean
Public Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean
Public Declare Function SetSystemTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME) As Boolean
Public Declare Function SetThreadIdealProcessor Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwIdealProcessor As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long, ByVal nPriority As Long) As Boolean
Public Declare Function SetUnhandledExceptionFilter Lib "kernel32.dll" (ByVal lpTopLevelExceptionFilter As Long) As Long
Public Declare Function SetVolumeLabel Lib "kernel32.dll" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Boolean
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Boolean
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
Public Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Boolean
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Boolean


Public Const CREATE_NEW As Long = 1
Public Const CREATE_ALWAYS As Long = 2
Public Const OPEN_EXISTING As Long = 3
Public Const OPEN_ALWAYS As Long = 4
Public Const TRUNCATE_EXISTING As Long = 5

Public Const DEBUG_PROCESS As Long = &H1
Public Const DEBUG_ONLY_THIS_PROCESS As Long = &H2
Public Const CREATE_SUSPENDED As Long = &H4
Public Const DETACHED_PROCESS As Long = &H8
Public Const CREATE_NEW_CONSOLE As Long = &H10
Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const IDLE_PRIORITY_CLASS As Long = &H40
Public Const HIGH_PRIORITY_CLASS As Long = &H80
Public Const REALTIME_PRIORITY_CLASS As Long = &H100
Public Const CREATE_NEW_PROCESS_GROUP As Long = &H200
Public Const CREATE_UNICODE_ENVIRONMENT As Long = &H400
Public Const CREATE_SEPARATE_WOW_VDM As Long = &H800
Public Const CREATE_SHARED_WOW_VDM As Long = &H1000
Public Const CREATE_FORCEDOS As Long = &H2000
Public Const BELOW_NORMAL_PRIORITY_CLASS As Long = &H4000
Public Const ABOVE_NORMAL_PRIORITY_CLASS As Long = &H8000
Public Const CREATE_BREAKAWAY_FROM_JOB As Long = &H1000000

Public Const WAIT_IO_COMPLETION As Long = STATUS_USER_APC
Public Const STILL_ACTIVE As Long = STATUS_PENDING
Public Const EXCEPTION_ACCESS_VIOLATION As Long = STATUS_ACCESS_VIOLATION
Public Const EXCEPTION_DATATYPE_MISALIGNMENT As Long = STATUS_DATATYPE_MISALIGNMENT
Public Const EXCEPTION_BREAKPOINT As Long = STATUS_BREAKPOINT
Public Const EXCEPTION_SINGLE_STEP As Long = STATUS_SINGLE_STEP
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED As Long = STATUS_ARRAY_BOUNDS_EXCEEDED
Public Const EXCEPTION_FLT_DENORMAL_OPERAND As Long = STATUS_FLOAT_DENORMAL_OPERAND
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO As Long = STATUS_FLOAT_DIVIDE_BY_ZERO
Public Const EXCEPTION_FLT_INEXACT_RESULT As Long = STATUS_FLOAT_INEXACT_RESULT
Public Const EXCEPTION_FLT_INVALID_OPERATION As Long = STATUS_FLOAT_INVALID_OPERATION
Public Const EXCEPTION_FLT_OVERFLOW As Long = STATUS_FLOAT_OVERFLOW
Public Const EXCEPTION_FLT_STACK_CHECK As Long = STATUS_FLOAT_STACK_CHECK
Public Const EXCEPTION_FLT_UNDERFLOW As Long = STATUS_FLOAT_UNDERFLOW
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO As Long = STATUS_INTEGER_DIVIDE_BY_ZERO
Public Const EXCEPTION_INT_OVERFLOW As Long = STATUS_INTEGER_OVERFLOW
Public Const EXCEPTION_PRIV_INSTRUCTION As Long = STATUS_PRIVILEGED_INSTRUCTION
Public Const EXCEPTION_IN_PAGE_ERROR As Long = STATUS_IN_PAGE_ERROR
Public Const EXCEPTION_ILLEGAL_INSTRUCTION As Long = STATUS_ILLEGAL_INSTRUCTION
Public Const EXCEPTION_NONCONTINUABLE_EXCEPTION As Long = STATUS_NONCONTINUABLE_EXCEPTION
Public Const EXCEPTION_STACK_OVERFLOW As Long = STATUS_STACK_OVERFLOW
Public Const EXCEPTION_INVALID_DISPOSITION As Long = STATUS_INVALID_DISPOSITION
Public Const EXCEPTION_GUARD_PAGE As Long = STATUS_GUARD_PAGE_VIOLATION
Public Const EXCEPTION_INVALID_HANDLE As Long = STATUS_INVALID_HANDLE
Public Const CONTROL_C_EXIT As Long = STATUS_CONTROL_C_EXIT

Public Const DONT_RESOLVE_DLL_REFERENCES As Long = &H1
Public Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2
Public Const LOAD_WITH_ALTERED_SEARCH_PATH As Long = &H8
Public Const LOAD_IGNORE_CODE_AUTHZ_LEVEL As Long = &H10

Public Const DOCKINFO_UNDOCKED As Long = &H1
Public Const DOCKINFO_DOCKED As Long = &H2
Public Const DOCKINFO_USER_SUPPLIED As Long = &H4
Public Const DOCKINFO_USER_UNDOCKED As Long = (DOCKINFO_USER_SUPPLIED Or DOCKINFO_UNDOCKED)
Public Const DOCKINFO_USER_DOCKED As Long = (DOCKINFO_USER_SUPPLIED Or DOCKINFO_DOCKED)

Public Const DRIVE_UNKNOWN As Long = 0
Public Const DRIVE_NO_ROOT_DIR As Long = 1
Public Const DRIVE_REMOVABLE As Long = 2
Public Const DRIVE_FIXED As Long = 3
Public Const DRIVE_REMOTE As Long = 4
Public Const DRIVE_CDROM As Long = 5
Public Const DRIVE_RAMDISK As Long = 6

Public Const FILE_BEGIN As Long = 0
Public Const FILE_CURRENT As Long = 1
Public Const FILE_END As Long = 2

Public Const FILE_ENCRYPTABLE As Long = 0
Public Const FILE_IS_ENCRYPTED As Long = 1
Public Const FILE_SYSTEM_ATTR As Long = 2
Public Const FILE_ROOT_DIR As Long = 3
Public Const FILE_SYSTEM_DIR As Long = 4
Public Const FILE_UNKNOWN As Long = 5
Public Const FILE_SYSTEM_NOT_SUPPORT As Long = 6
Public Const FILE_USER_DISALLOWED As Long = 7
Public Const FILE_READ_ONLY As Long = 8
Public Const FILE_DIR_DISALLOWED As Long = 9

Public Const FILE_TYPE_UNKNOWN As Long = &H0
Public Const FILE_TYPE_DISK As Long = &H1
Public Const FILE_TYPE_CHAR As Long = &H2
Public Const FILE_TYPE_PIPE As Long = &H3
Public Const FILE_TYPE_REMOTE As Long = &H8000

Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Public Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Public Const FORMAT_MESSAGE_FROM_STRING As Long = &H400
Public Const FORMAT_MESSAGE_FROM_HMODULE As Long = &H800
Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = &H2000
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = &HFF

Public Const FS_CASE_IS_PRESERVED As Long = FILE_CASE_PRESERVED_NAMES
Public Const FS_CASE_SENSITIVE As Long = FILE_CASE_SENSITIVE_SEARCH
Public Const FS_UNICODE_STORED_ON_DISK As Long = FILE_UNICODE_ON_DISK
Public Const FS_PERSISTENT_ACLS As Long = FILE_PERSISTENT_ACLS
Public Const FS_VOL_IS_COMPRESSED As Long = FILE_VOLUME_IS_COMPRESSED
Public Const FS_FILE_COMPRESSION As Long = FILE_FILE_COMPRESSION
Public Const FS_FILE_ENCRYPTION As Long = FILE_SUPPORTS_ENCRYPTION

Public Const HW_PROFILE_GUIDLEN As Long = 39

Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const INVALID_FILE_SIZE As Long = &HFFFFFFFF
Public Const INVALID_SET_FILE_POINTER As Long = -1

Public Const MAX_PROFILE_LEN As Long = 80

Public Const MAX_COMPUTERNAME_LENGTH As Long = 31

Public Const MAXLONG As Long = &H7FFFFFFF

Public Const SEM_FAILCRITICALERRORS As Long = &H1
Public Const SEM_NOGPFAULTERRORBOX As Long = &H2
Public Const SEM_NOALIGNMENTFAULTEXCEPT As Long = &H4
Public Const SEM_NOOPENFILEERRORBOX As Long = &H8000

Public Const THREAD_PRIORITY_LOWEST As Long = THREAD_BASE_PRIORITY_MIN
Public Const THREAD_PRIORITY_BELOW_NORMAL As Long = (THREAD_PRIORITY_LOWEST + 1)
Public Const THREAD_PRIORITY_NORMAL As Long = 0
Public Const THREAD_PRIORITY_HIGHEST As Long = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_ABOVE_NORMAL As Long = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_ERROR_RETURN As Long = (MAXLONG)
Public Const THREAD_PRIORITY_TIME_CRITICAL As Long = THREAD_BASE_PRIORITY_LOWRT
Public Const THREAD_PRIORITY_IDLE As Long = THREAD_BASE_PRIORITY_IDLE

Public Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF


Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Public Type HW_PROFILE_INFO
    dwDockInfo As Long
    szHwProfileGuid As String * HW_PROFILE_GUIDLEN
    szHwProfileName As String * MAX_PROFILE_LEN
End Type

Public Type SYSTEM_INFO
    dwOemID As Long 'Union
    'WORD wProcessorArchitecture
    'WORD wReserved

    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Public Type SYSTEM_POWER_STATUS
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    Reserved1 As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName As String * 64
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName As String * 64
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
 
Public Enum COMPUTER_NAME_FORMAT
    ComputerNameNetBIOS
    ComputerNameDnsHostname
    ComputerNameDnsDomain
    ComputerNameDnsFullyQualified
    ComputerNamePhysicalNetBIOS
    ComputerNamePhysicalDnsHostname
    ComputerNamePhysicalDnsDomain
    ComputerNamePhysicalDnsFullyQualified
    ComputerNameMax
End Enum
