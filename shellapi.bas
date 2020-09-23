Attribute VB_Name = "shellapi"
Option Explicit


Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, ByRef pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Boolean
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long


Public Const NIM_ADD As Long = &H0
Public Const NIM_MODIFY As Long = &H1
Public Const NIM_DELETE As Long = &H2
Public Const NIM_SETFOCUS As Long = &H3
Public Const NIM_SETVERSION As Long = &H4

Public Const NIF_MESSAGE As Long = &H1
Public Const NIF_ICON As Long = &H2
Public Const NIF_TIP As Long = &H4
Public Const NIF_STATE As Long = &H8
Public Const NIF_INFO As Long = &H10

Public Const SEE_MASK_CLASSNAME As Long = &H1
Public Const SEE_MASK_CLASSKEY As Long = &H3
Public Const SEE_MASK_IDLIST As Long = &H4
Public Const SEE_MASK_INVOKEIDLIST As Long = &HC
Public Const SEE_MASK_ICON As Long = &H10
Public Const SEE_MASK_HOTKEY As Long = &H20
Public Const SEE_MASK_NOCLOSEPROCESS As Long = &H40
Public Const SEE_MASK_CONNECTNETDRV As Long = &H80
Public Const SEE_MASK_FLAG_DDEWAIT As Long = &H100
Public Const SEE_MASK_DOENVSUBST As Long = &H200
Public Const SEE_MASK_FLAG_NO_UI As Long = &H400
Public Const SEE_MASK_UNICODE As Long = &H4000
Public Const SEE_MASK_NO_CONSOLE As Long = &H8000
Public Const SEE_MASK_ASYNCOK As Long = &H100000
Public Const SEE_MASK_HMONITOR As Long = &H200000
Public Const SEE_MASK_NOQUERYCLASSSTORE As Long = &H1000000
Public Const SEE_MASK_WAITFORINPUTIDLE As Long = &H2000000
Public Const SEE_MASK_FLAG_LOG_USAGE As Long = &H4000000

Public Const SHERB_NOCONFIRMATION As Long = &H1
Public Const SHERB_NOPROGRESSUI As Long = &H2
Public Const SHERB_NOSOUND As Long = &H4


Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
 
    'Optional members
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    
    'Union
    uhIconMonitor As Long
    'HANDLE hIcon
    'HANDLE hMonitor
    
    hProcess As Long
End Type
