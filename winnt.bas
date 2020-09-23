Attribute VB_Name = "winnt"
Option Explicit


Public Const DELETE As Long = &H10000
Public Const READ_CONTROL As Long = &H20000
Public Const WRITE_DAC As Long = &H40000
Public Const WRITE_OWNER As Long = &H80000
Public Const SYNCHRONIZE As Long = &H100000
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL As Long = &HFFFF
Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const GENERIC_EXECUTE As Long = &H20000000
Public Const GENERIC_ALL As Long = &H10000000

Public Const STATUS_WAIT_0 As Long = &H0
Public Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Public Const STATUS_USER_APC As Long = &HC0
Public Const STATUS_TIMEOUT As Long = &H102
Public Const STATUS_PENDING As Long = &H103
Public Const DBG_EXCEPTION_HANDLED As Long = &H10001
Public Const DBG_CONTINUE As Long = &H10002
Public Const STATUS_SEGMENT_NOTIFICATION As Long = &H40000005
Public Const DBG_TERMINATE_THREAD As Long = &H40010003
Public Const DBG_TERMINATE_PROCESS As Long = &H40010004
Public Const DBG_CONTROL_C As Long = &H40010005
Public Const DBG_CONTROL_BREAK As Long = &H40010008
Public Const STATUS_GUARD_PAGE_VIOLATION As Long = &H80000001
Public Const STATUS_DATATYPE_MISALIGNMENT As Long = &H80000002
Public Const STATUS_BREAKPOINT As Long = &H80000003
Public Const STATUS_SINGLE_STEP As Long = &H80000004
Public Const DBG_EXCEPTION_NOT_HANDLED As Long = &H80010001
Public Const STATUS_ACCESS_VIOLATION As Long = &HC0000005
Public Const STATUS_IN_PAGE_ERROR As Long = &HC0000006
Public Const STATUS_INVALID_HANDLE As Long = &HC0000008
Public Const STATUS_NO_MEMORY As Long = &HC0000017
Public Const STATUS_ILLEGAL_INSTRUCTION As Long = &HC000001D
Public Const STATUS_NONCONTINUABLE_EXCEPTION As Long = &HC0000025
Public Const STATUS_INVALID_DISPOSITION As Long = &HC0000026
Public Const STATUS_ARRAY_BOUNDS_EXCEEDED As Long = &HC000008C
Public Const STATUS_FLOAT_DENORMAL_OPERAND As Long = &HC000008D
Public Const STATUS_FLOAT_DIVIDE_BY_ZERO As Long = &HC000008E
Public Const STATUS_FLOAT_INEXACT_RESULT As Long = &HC000008F
Public Const STATUS_FLOAT_INVALID_OPERATION As Long = &HC0000090
Public Const STATUS_FLOAT_OVERFLOW As Long = &HC0000091
Public Const STATUS_FLOAT_STACK_CHECK As Long = &HC0000092
Public Const STATUS_FLOAT_UNDERFLOW As Long = &HC0000093
Public Const STATUS_INTEGER_DIVIDE_BY_ZERO As Long = &HC0000094
Public Const STATUS_INTEGER_OVERFLOW As Long = &HC0000095
Public Const STATUS_PRIVILEGED_INSTRUCTION As Long = &HC0000096
Public Const STATUS_STACK_OVERFLOW As Long = &HC00000FD
Public Const STATUS_CONTROL_C_EXIT As Long = &HC000013A
Public Const STATUS_FLOAT_MULTIPLE_FAULTS As Long = &HC00002B4
Public Const STATUS_FLOAT_MULTIPLE_TRAPS As Long = &HC00002B5
Public Const STATUS_REG_NAT_CONSUMPTION As Long = &HC00002C9
Public Const STATUS_SXS_EARLY_DEACTIVATION As Long = &HC015000F
Public Const STATUS_SXS_INVALID_DEACTIVATION As Long = &HC0150010

Public Const EXCEPTION_NONCONTINUABLE As Long = &H1
Public Const EXCEPTION_MAXIMUM_PARAMETERS As Long = 15

Public Const FILE_CASE_SENSITIVE_SEARCH As Long = &H1
Public Const FILE_CASE_PRESERVED_NAMES As Long = &H2
Public Const FILE_UNICODE_ON_DISK As Long = &H4
Public Const FILE_PERSISTENT_ACLS As Long = &H8
Public Const FILE_FILE_COMPRESSION As Long = &H10
Public Const FILE_VOLUME_QUOTAS As Long = &H20
Public Const FILE_SUPPORTS_SPARSE_FILES As Long = &H40
Public Const FILE_SUPPORTS_REPARSE_POINTS As Long = &H80
Public Const FILE_SUPPORTS_REMOTE_STORAGE As Long = &H100
Public Const FILE_VOLUME_IS_COMPRESSED As Long = &H8000
Public Const FILE_SUPPORTS_OBJECT_IDS As Long = &H10000
Public Const FILE_SUPPORTS_ENCRYPTION As Long = &H20000
Public Const FILE_NAMED_STREAMS As Long = &H40000

Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const FILE_ATTRIBUTE_DEVICE As Long = &H40
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Public Const FILE_ATTRIBUTE_SPARSE_FILE As Long = &H200
Public Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400
Public Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Public Const FILE_ATTRIBUTE_OFFLINE As Long = &H1000
Public Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED As Long = &H2000
Public Const FILE_ATTRIBUTE_ENCRYPTED As Long = &H4000

Public Const FILE_SHARE_READ As Long = &H1
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const FILE_SHARE_DELETE As Long = &H4

Public Const IMAGE_DOS_SIGNATURE As Long = &H4D5A
Public Const IMAGE_OS2_SIGNATURE As Long = &H4E45
Public Const IMAGE_OS2_SIGNATURE_LE As Long = &H4C45
Public Const IMAGE_NT_SIGNATURE As Long = &H50450000

Public Const IMAGE_SIZEOF_FILE_HEADER As Long = 20

Public Const IMAGE_FILE_RELOCS_STRIPPED As Long = &H1
Public Const IMAGE_FILE_EXECUTABLE_IMAGE As Long = &H2
Public Const IMAGE_FILE_LINE_NUMS_STRIPPED As Long = &H4
Public Const IMAGE_FILE_LOCAL_SYMS_STRIPPED As Long = &H8
Public Const IMAGE_FILE_AGGRESIVE_WS_TRIM As Long = &H10
Public Const IMAGE_FILE_LARGE_ADDRESS_AWARE As Long = &H20
Public Const IMAGE_FILE_BYTES_REVERSED_LO As Long = &H80
Public Const IMAGE_FILE_32BIT_MACHINE As Long = &H100
Public Const IMAGE_FILE_DEBUG_STRIPPED As Long = &H200
Public Const IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP As Long = &H400
Public Const IMAGE_FILE_NET_RUN_FROM_SWAP As Long = &H800
Public Const IMAGE_FILE_SYSTEM As Long = &H1000
Public Const IMAGE_FILE_DLL As Long = &H2000
Public Const IMAGE_FILE_UP_SYSTEM_ONLY As Long = &H4000
Public Const IMAGE_FILE_BYTES_REVERSED_HI As Long = &H8000

Public Const IMAGE_FILE_MACHINE_UNKNOWN As Long = 0
Public Const IMAGE_FILE_MACHINE_I386 As Long = &H14C
Public Const IMAGE_FILE_MACHINE_R3000 As Long = &H162
Public Const IMAGE_FILE_MACHINE_R4000 As Long = &H166
Public Const IMAGE_FILE_MACHINE_R10000 As Long = &H168
Public Const IMAGE_FILE_MACHINE_WCEMIPSV2 As Long = &H169
Public Const IMAGE_FILE_MACHINE_ALPHA As Long = &H184
Public Const IMAGE_FILE_MACHINE_SH3 As Long = &H1A2
Public Const IMAGE_FILE_MACHINE_SH3DSP As Long = &H1A3
Public Const IMAGE_FILE_MACHINE_SH3E As Long = &H1A4
Public Const IMAGE_FILE_MACHINE_SH4 As Long = &H1A6
Public Const IMAGE_FILE_MACHINE_SH5 As Long = &H1A8
Public Const IMAGE_FILE_MACHINE_ARM As Long = &H1C0
Public Const IMAGE_FILE_MACHINE_THUMB As Long = &H1C2
Public Const IMAGE_FILE_MACHINE_AM33 As Long = &H1D3
Public Const IMAGE_FILE_MACHINE_POWERPC As Long = &H1F0
Public Const IMAGE_FILE_MACHINE_POWERPCFP As Long = &H1F1
Public Const IMAGE_FILE_MACHINE_IA64 As Long = &H200
Public Const IMAGE_FILE_MACHINE_MIPS16 As Long = &H266
Public Const IMAGE_FILE_MACHINE_ALPHA64 As Long = &H284
Public Const IMAGE_FILE_MACHINE_MIPSFPU As Long = &H366
Public Const IMAGE_FILE_MACHINE_MIPSFPU16 As Long = &H466
Public Const IMAGE_FILE_MACHINE_AXP64 As Long = IMAGE_FILE_MACHINE_ALPHA64
Public Const IMAGE_FILE_MACHINE_TRICORE As Long = &H520
Public Const IMAGE_FILE_MACHINE_CEF As Long = &HCEF
Public Const IMAGE_FILE_MACHINE_EBC As Long = &HEBC
Public Const IMAGE_FILE_MACHINE_AMD64 As Long = &H8664
Public Const IMAGE_FILE_MACHINE_M32R As Long = &H9041
Public Const IMAGE_FILE_MACHINE_CEE As Long = &HC0EE

Public Const KEY_CREATE_LINK As Long = &H20
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_EVENT As Long = &H1
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_WRITE As Long = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE As Long = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_WOW64_32KEY As Long = &H200
Public Const KEY_WOW64_64KEY As Long = &H100
Public Const KEY_WOW64_RES As Long = &H300

Public Const MAXIMUM_PROCESSORS As Long = 32

Public Const PROCESS_TERMINATE As Long = &H1
Public Const PROCESS_CREATE_THREAD As Long = &H2
Public Const PROCESS_SET_SESSIONID As Long = &H4
Public Const PROCESS_VM_OPERATION As Long = &H8
Public Const PROCESS_VM_READ As Long = &H10
Public Const PROCESS_VM_WRITE As Long = &H20
Public Const PROCESS_DUP_HANDLE As Long = &H40
Public Const PROCESS_CREATE_PROCESS As Long = &H80
Public Const PROCESS_SET_QUOTA As Long = &H100
Public Const PROCESS_SET_INFORMATION As Long = &H200
Public Const PROCESS_QUERY_INFORMATION As Long = &H400
Public Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

Public Const PROCESSOR_ARCHITECTURE_INTEL As Long = 0
Public Const PROCESSOR_ARCHITECTURE_MIPS As Long = 1
Public Const PROCESSOR_ARCHITECTURE_ALPHA As Long = 2
Public Const PROCESSOR_ARCHITECTURE_PPC As Long = 3
Public Const PROCESSOR_ARCHITECTURE_SHX As Long = 4
Public Const PROCESSOR_ARCHITECTURE_ARM As Long = 5
Public Const PROCESSOR_ARCHITECTURE_IA64 As Long = 6
Public Const PROCESSOR_ARCHITECTURE_ALPHA64 As Long = 7
Public Const PROCESSOR_ARCHITECTURE_MSIL As Long = 8
Public Const PROCESSOR_ARCHITECTURE_AMD64 As Long = 9
Public Const PROCESSOR_ARCHITECTURE_IA32_ON_WIN64 As Long = 10
Public Const PROCESSOR_ARCHITECTURE_UNKNOWN As Long = &HFFFF

Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_DWORD_LITTLE_ENDIAN As Long = 4
Public Const REG_DWORD_BIG_ENDIAN As Long = 5
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_RESOURCE_LIST As Long = 8
Public Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9
Public Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10
Public Const REG_QWORD As Long = 11
Public Const REG_QWORD_LITTLE_ENDIAN As Long = 11

Public Const REG_OPTION_RESERVED As Long = 0
Public Const REG_OPTION_NON_VOLATILE As Long = 0
Public Const REG_OPTION_VOLATILE As Long = 1
Public Const REG_OPTION_CREATE_LINK As Long = 2
Public Const REG_OPTION_BACKUP_RESTORE As Long = 4
Public Const REG_OPTION_OPEN_LINK As Long = &H8
Public Const REG_LEGAL_OPTION As Long = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE Or REG_OPTION_OPEN_LINK)

Public Const SE_CREATE_TOKEN_NAME As String = "SeCreateTokenPrivilege"
Public Const SE_ASSIGNPRIMARYTOKEN_NAME As String = "SeAssignPrimaryTokenPrivilege"
Public Const SE_LOCK_MEMORY_NAME As String = "SeLockMemoryPrivilege"
Public Const SE_INCREASE_QUOTA_NAME As String = "SeIncreaseQuotaPrivilege"
Public Const SE_UNSOLICITED_INPUT_NAME As String = "SeUnsolicitedInputPrivilege"
Public Const SE_MACHINE_ACCOUNT_NAME As String = "SeMachineAccountPrivilege"
Public Const SE_TCB_NAME As String = "SeTcbPrivilege"
Public Const SE_SECURITY_NAME As String = "SeSecurityPrivilege"
Public Const SE_TAKE_OWNERSHIP_NAME As String = "SeTakeOwnershipPrivilege"
Public Const SE_LOAD_DRIVER_NAME As String = "SeLoadDriverPrivilege"
Public Const SE_SYSTEM_PROFILE_NAME As String = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME_NAME As String = "SeSystemtimePrivilege"
Public Const SE_PROF_SINGLE_PROCESS_NAME As String = "SeProfileSingleProcessPrivilege"
Public Const SE_INC_BASE_PRIORITY_NAME As String = "SeIncreaseBasePriorityPrivilege"
Public Const SE_CREATE_PAGEFILE_NAME As String = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT_NAME As String = "SeCreatePermanentPrivilege"
Public Const SE_BACKUP_NAME As String = "SeBackupPrivilege"
Public Const SE_RESTORE_NAME As String = "SeRestorePrivilege"
Public Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"
Public Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
Public Const SE_AUDIT_NAME As String = "SeAuditPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT_NAME As String = "SeSystemEnvironmentPrivilege"
Public Const SE_CHANGE_NOTIFY_NAME As String = "SeChangeNotifyPrivilege"
Public Const SE_REMOTE_SHUTDOWN_NAME As String = "SeRemoteShutdownPrivilege"
Public Const SE_UNDOCK_NAME As String = "SeUndockPrivilege"
Public Const SE_SYNC_AGENT_NAME As String = "SeSyncAgentPrivilege"
Public Const SE_ENABLE_DELEGATION_NAME As String = "SeEnableDelegationPrivilege"
Public Const SE_MANAGE_VOLUME_NAME As String = "SeManageVolumePrivilege"

Public Const SE_PRIVILEGE_ENABLED_BY_DEFAULT As Long = &H1
Public Const SE_PRIVILEGE_ENABLED As Long = &H2
Public Const SE_PRIVILEGE_USED_FOR_ACCESS As Long = &H80000000

Public Const THREAD_TERMINATE As Long = &H1
Public Const THREAD_SUSPEND_RESUME As Long = &H2
Public Const THREAD_GET_CONTEXT As Long = &H8
Public Const THREAD_SET_CONTEXT As Long = &H10
Public Const THREAD_SET_INFORMATION As Long = &H20
Public Const THREAD_QUERY_INFORMATION As Long = &H40
Public Const THREAD_SET_THREAD_TOKEN As Long = &H80
Public Const THREAD_IMPERSONATE As Long = &H100
Public Const THREAD_DIRECT_IMPERSONATION As Long = &H200
Public Const THREAD_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)

Public Const THREAD_BASE_PRIORITY_LOWRT As Long = 15
Public Const THREAD_BASE_PRIORITY_MAX As Long = 2
Public Const THREAD_BASE_PRIORITY_MIN As Long = -2
Public Const THREAD_BASE_PRIORITY_IDLE As Long = -15

Public Const TIME_ZONE_ID_UNKNOWN As Long = 0
Public Const TIME_ZONE_ID_STANDARD As Long = 1
Public Const TIME_ZONE_ID_DAYLIGHT As Long = 2

Public Const TOKEN_ASSIGN_PRIMARY As Long = &H1
Public Const TOKEN_DUPLICATE As Long = &H2
Public Const TOKEN_IMPERSONATE As Long = &H4
Public Const TOKEN_QUERY As Long = &H8
Public Const TOKEN_QUERY_SOURCE As Long = &H10
Public Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Public Const TOKEN_ADJUST_GROUPS As Long = &H40
Public Const TOKEN_ADJUST_DEFAULT As Long = &H80
Public Const TOKEN_ADJUST_SESSIONID As Long = &H100
Public Const TOKEN_ALL_ACCESS_P As Long = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
Public Const TOKEN_ALL_ACCESS_NT As Long = (TOKEN_ALL_ACCESS_P Or TOKEN_ADJUST_SESSIONID)
Public Const TOKEN_ALL_ACCESS As Long = (TOKEN_ALL_ACCESS_P)
Public Const TOKEN_READ As Long = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)
Public Const TOKEN_WRITE As Long = (STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
Public Const TOKEN_EXECUTE As Long = STANDARD_RIGHTS_EXECUTE

Public Const VER_NT_WORKSTATION As Long = &H1
Public Const VER_NT_DOMAIN_CONTROLLER As Long = &H2
Public Const VER_NT_SERVER As Long = &H3

Public Const VER_PLATFORM_WIN32s As Long = 0
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2

Public Const VER_SUITE_SMALLBUSINESS As Long = &H1
Public Const VER_SUITE_ENTERPRISE As Long = &H2
Public Const VER_SUITE_BACKOFFICE As Long = &H4
Public Const VER_SUITE_COMMUNICATIONS As Long = &H8
Public Const VER_SUITE_TERMINAL As Long = &H10
Public Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20
Public Const VER_SUITE_EMBEDDEDNT As Long = &H40
Public Const VER_SUITE_DATACENTER As Long = &H80
Public Const VER_SUITE_SINGLEUSERTS As Long = &H100


Public Type ADMINISTRATOR_POWER_POLICY
    MinSleep As SYSTEM_POWER_STATE
    MaxSleep As SYSTEM_POWER_STATE
    MinVideoTimeout As Long
    MaxVideoTimeout As Long
    MinSpindownTimeout As Long
    MaxSpindownTimeout As Long
End Type

Public Type BATTERY_REPORTING_SCALE
    Granularity As Long
    Capacity As Long
End Type

Public Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    ExceptionRecord As Long
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(0 To EXCEPTION_MAXIMUM_PARAMETERS - 1) As Long
End Type

Public Type EXCEPTION_POINTERS
    ExceptionRecord As EXCEPTION_RECORD
    ContextRecord As Long
End Type

Public Type IMAGE_DOS_HEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(0 To 3)   As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(0 To 9) As Integer
    e_lfanew As Long
End Type

Public Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Public Type IMAGE_OS2_HEADER
    ne_magic As Integer
    ne_ver As Byte
    ne_rev As Byte
    ne_enttab As Integer
    ne_cbenttab As Integer
    ne_crc As Long
    ne_flags As Integer
    ne_autodata As Integer
    ne_heap As Integer
    ne_stack As Integer
    ne_csip As Long
    ne_sssp As Long
    ne_cseg As Integer
    ne_cmod As Integer
    ne_cbnrestab As Integer
    ne_segtab As Integer
    ne_rsrctab As Integer
    ne_restab As Integer
    ne_modtab As Integer
    ne_imptab As Integer
    ne_nrestab As Long
    ne_cmovent As Integer
    ne_align As Integer
    ne_cres As Integer
    ne_exetyp As Byte
    ne_flagsothers As Byte
    ne_pretthunks As Integer
    ne_psegrefbytes As Integer
    ne_swaparea As Integer
    ne_expver As Integer
End Type
  
Public Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Public Type LUID
    LowPart As Long
    HighPart As Long
End Type

Public Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Public Type IO_COUNTERS
    ReadOperationCount As LARGE_INTEGER
    WriteOperationCount As LARGE_INTEGER
    OtherOperationCount As LARGE_INTEGER
    ReadTransferCount As LARGE_INTEGER
    WriteTransferCount As LARGE_INTEGER
    OtherTransferCount As LARGE_INTEGER
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Public Type PROCESSOR_POWER_INFORMATION
    Number As Long
    MaxMhz As Long
    CurrentMhz As Long
    MhzLimit As Long
    MaxIdleState As Long
    CurrentIdleState As Long
End Type

Public Type SYSTEM_BATTERY_STATE
    AcOnLine As Byte
    BatteryPresent As Byte
    Charging As Byte
    Discharging As Byte
    Spare1(0 To 3) As Byte
    MaxCapacity As Long
    RemainingCapacity As Long
    Rate As Long
    EstimatedTime As Long
    DefaultAlert1 As Long
    DefaultAlert2 As Long
End Type

Public Type SYSTEM_POWER_CAPABILITIES
    PowerButtonPresent As Byte
    SleepButtonPresent As Byte
    LidPresent As Byte
    SystemS1 As Byte
    SystemS2 As Byte
    SystemS3 As Byte
    SystemS4 As Byte
    SystemS5 As Byte
    HiberFilePresent As Byte
    FullWake As Byte
    VideoDimPresent As Byte
    ApmPresent As Byte
    UpsPresent As Byte
    ThermalControl As Byte
    ProcessorThrottle As Byte
    ProcessorMinThrottle As Byte
    ProcessorMaxThrottle As Byte
    spare2(0 To 3) As Byte
    DiskSpinDown As Byte
    spare3(0 To 7) As Byte
    SystemBatteriesPresent As Byte
    BatteriesAreShortTerm As Byte
    BatteryScale(0 To 2) As BATTERY_REPORTING_SCALE
    AcOnLineWake As SYSTEM_POWER_STATE
    SoftLidWake As SYSTEM_POWER_STATE
    RtcWake As SYSTEM_POWER_STATE
    MinDeviceWakeState As SYSTEM_POWER_STATE
    DefaultLowLatencyWake As SYSTEM_POWER_STATE
End Type

Public Type SYSTEM_POWER_INFORMATION
    MaxIdlenessAllowed As Long
    Idleness As Long
    TimeRemaining As Long
    CoolingMode As Long
End Type

Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(64) As LUID_AND_ATTRIBUTES
End Type


Public Enum POWER_INFORMATION_LEVEL
    SystemPowerPolicyAc = 0
    SystemPowerPolicyDc = 1
    VerifySystemPolicyAc = 2
    VerifySystemPolicyDc = 3
    SystemPowerCapabilities = 4
    SystemBatteryState = 5
    SystemPowerStateHandler = 6
    ProcessorStateHandler = 7
    SystemPowerPolicyCurrent = 8
    AdministratorPowerPolicy = 9
    SystemReserveHiberFile = 10
    ProcessorInformation = 11
    SystemPowerInformation = 12
    ProcessorStateHandler2 = 13
    LastWakeTime = 14
    LastSleepTime = 15
    SystemExecutionState = 16
    SystemPowerStateNotifyHandler = 17
    ProcessorPowerPolicyAc = 18
    ProcessorPowerPolicyDc = 19
    VerifyProcessorPowerPolicyAc = 20
    VerifyProcessorPowerPolicyDc = 21
    ProcessorPowerPolicyCurrent = 22
End Enum

Public Enum SYSTEM_POWER_STATE
    PowerSystemUnspecified = 0
    PowerSystemWorking = 1
    PowerSystemSleeping1 = 2
    PowerSystemSleeping2 = 3
    PowerSystemSleeping3 = 4
    PowerSystemHibernate = 5
    PowerSystemShutdown = 6
    PowerSystemMaximum = 7
End Enum
