Attribute VB_Name = "winreg"
Option Explicit


Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As Any, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByRef lpftLastWriteTime As Any) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByVal lpReserved As Long, ByRef lpcSubKeys As Long, ByRef lpcbMaxSubKeyLen As Long, ByRef lpcbMaxClassLen As Long, ByRef lpcValues As Long, ByRef lpcbMaxValueNameLen As Long, ByRef lpcbMaxValueLen As Long, ByRef lpcbSecurityDescriptor As Long, ByRef lpftLastWriteTime As Any) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long


Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const HKEY_DYN_DATA As Long = &H80000006
