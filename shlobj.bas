Attribute VB_Name = "shlobj"
Option Explicit


Public Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Boolean) As Boolean
Public Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long


Public Const CSIDL_DESKTOP As Long = &H0
Public Const CSIDL_INTERNET As Long = &H1
Public Const CSIDL_PROGRAMS As Long = &H2
Public Const CSIDL_CONTROLS As Long = &H3
Public Const CSIDL_PRINTERS As Long = &H4
Public Const CSIDL_PERSONAL As Long = &H5
Public Const CSIDL_FAVORITES As Long = &H6
Public Const CSIDL_STARTUP As Long = &H7
Public Const CSIDL_RECENT As Long = &H8
Public Const CSIDL_SENDTO As Long = &H9
Public Const CSIDL_BITBUCKET As Long = &HA
Public Const CSIDL_STARTMENU As Long = &HB
Public Const CSIDL_DESKTOPDIRECTORY As Long = &H10
Public Const CSIDL_DRIVES As Long = &H11
Public Const CSIDL_NETWORK As Long = &H12
Public Const CSIDL_NETHOOD As Long = &H13
Public Const CSIDL_FONTS As Long = &H14
Public Const CSIDL_TEMPLATES As Long = &H15
Public Const CSIDL_COMMON_STARTMENU As Long = &H16
Public Const CSIDL_COMMON_PROGRAMS As Long = &H17
Public Const CSIDL_COMMON_STARTUP As Long = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19
Public Const CSIDL_APPDATA As Long = &H1A
Public Const CSIDL_PRINTHOOD As Long = &H1B
Public Const CSIDL_LOCAL_APPDATA As Long = &H1C
Public Const CSIDL_ALTSTARTUP As Long = &H1D
Public Const CSIDL_COMMON_ALTSTARTUP As Long = &H1E
Public Const CSIDL_COMMON_FAVORITES As Long = &H1F
Public Const CSIDL_INTERNET_CACHE As Long = &H20
Public Const CSIDL_COOKIES As Long = &H21
Public Const CSIDL_HISTORY As Long = &H22
Public Const CSIDL_COMMON_APPDATA As Long = &H23
Public Const CSIDL_WINDOWS As Long = &H24
Public Const CSIDL_SYSTEM As Long = &H25
Public Const CSIDL_PROGRAM_FILES As Long = &H26
Public Const CSIDL_MYPICTURES As Long = &H27
Public Const CSIDL_PROFILE As Long = &H28
Public Const CSIDL_SYSTEMX86 As Long = &H29
Public Const CSIDL_PROGRAM_FILESX86 As Long = &H2A
Public Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B
Public Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C
Public Const CSIDL_COMMON_TEMPLATES As Long = &H2D
Public Const CSIDL_COMMON_DOCUMENTS As Long = &H2E
Public Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F
Public Const CSIDL_ADMINTOOLS As Long = &H30
Public Const CSIDL_CONNECTIONS As Long = &H31

Public Const CSIDL_FLAG_CREATE As Long = &H8000
Public Const CSIDL_FLAG_DONT_VERIFY As Long = &H4000
Public Const CSIDL_FLAG_MASK As Long = &HFF00


Public Enum SHGFP_TYPE
    SHGFP_TYPE_CURRENT = 0&
    SHGFP_TYPE_DEFAULT = 1&
End Enum
