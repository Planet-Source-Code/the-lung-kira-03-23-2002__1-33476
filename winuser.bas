Attribute VB_Name = "winuser"
Option Explicit


Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ChangeDisplaySettings Lib "user32.dll" Alias "ChangeDisplaySettingsA" (ByRef lpDevMode As DEVMODE, ByVal dwFlags As Long) As Long
Public Declare Function CharUpper Lib "user32.dll" Alias "CharUpperA" (ByVal lpsz As String) As String
Public Declare Function EnumDisplayDevices Lib "user32.dll" Alias "EnumDisplayDevicesA" (ByVal lpDevice As Long, ByVal iDevNum As Long, ByRef lpDisplayDevice As DISPLAY_DEVICE, ByVal dwFlags As Long) As Boolean
Public Declare Function EnumDisplayMonitors Lib "user32.dll" (ByVal hdc As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Boolean
Public Declare Function EnumDisplaySettings Lib "user32.dll" Alias "EnumDisplaySettingsA" (ByRef lpszDeviceName As Any, ByVal iModeNum As Long, ByRef lpDevMode As DEVMODE) As Boolean
Public Declare Function EnumDisplaySettingsEx Lib "user32.dll" Alias "EnumDisplaySettingsExA" (ByRef lpszDeviceName As Any, ByVal iModeNum As Long, ByRef lpDevMode As DEVMODE, ByVal dwFlags As Long) As Boolean
Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
Public Declare Function FlashWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal bInvert As Long) As Boolean
Public Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Public Declare Function GetCaretBlinkTime Lib "user32.dll" () As Long
Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Boolean
Public Declare Function GetDoubleClickTime Lib "user32.dll" () As Long
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function GetGuiResources Lib "user32.dll" (ByVal hProcess As Long, ByVal uiFlags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32.dll" (ByVal idThread As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32.dll" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Boolean
Public Declare Function GetKeyboardType Lib "user32.dll" (ByVal nTypeFlag As Long) As Long
Public Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFOEX) As Boolean
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal uCmd As Long) As Long
Public Declare Function GetWindowInfo Lib "user32.dll" (ByVal hwnd As Long, ByRef pwi As WINDOWINFO) As Boolean
Public Declare Function GetWindowModuleFileName Lib "user32.dll" Alias "GetWindowModuleFileNameA" (ByVal hwnd As Long, ByVal lpszFileName As String, ByVal cchFileNameMax As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUS)
Public Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUSEX) As Boolean
Public Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Boolean
Public Declare Function LockWorkStation Lib "user32.dll" () As Boolean
Public Declare Function MessageBoxEx Lib "user32.dll" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Integer) As Long
Public Declare Function SendMessage Lib "user32.dll" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetCaretBlinkTime Lib "user32.dll" (ByVal wMSeconds As Long) As Boolean
Public Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Boolean
Public Declare Function SetDoubleClickTime Lib "user32.dll" (ByVal wCount As Long) As Boolean
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Boolean
Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Boolean
Public Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uiAction As Long, ByVal uiParam As Long, ByRef pvParam As Any, ByVal fWinIni As Long) As Boolean


Public Const ARW_BOTTOMLEFT As Long = &H0
Public Const ARW_BOTTOMRIGHT As Long = &H1
Public Const ARW_TOPLEFT As Long = &H2
Public Const ARW_TOPRIGHT As Long = &H3
Public Const ARW_STARTMASK As Long = &H3
Public Const ARW_STARTRIGHT As Long = &H1
Public Const ARW_STARTTOP As Long = &H2

Public Const ARW_LEFT As Long = &H0
Public Const ARW_RIGHT As Long = &H0
Public Const ARW_UP As Long = &H4
Public Const ARW_DOWN As Long = &H4
Public Const ARW_HIDE As Long = &H8

Public Const ATF_TIMEOUTON As Long = &H1
Public Const ATF_ONOFFFEEDBACK As Long = &H2

Public Const CDS_UPDATEREGISTRY As Long = &H1
Public Const CDS_TEST As Long = &H2
Public Const CDS_FULLSCREEN As Long = &H4
Public Const CDS_GLOBAL As Long = &H8
Public Const CDS_SET_PRIMARY As Long = &H10
Public Const CDS_VIDEOPARAMETERS As Long = &H20
Public Const CDS_RESET As Long = &H40000000
Public Const CDS_NORESET As Long = &H10000000

Public Const DISP_CHANGE_SUCCESSFUL As Long = 0
Public Const DISP_CHANGE_RESTART As Long = 1
Public Const DISP_CHANGE_FAILED As Long = -1
Public Const DISP_CHANGE_BADMODE As Long = -2
Public Const DISP_CHANGE_NOTUPDATED As Long = -3
Public Const DISP_CHANGE_BADFLAGS As Long = -4
Public Const DISP_CHANGE_BADPARAM As Long = -5
Public Const DISP_CHANGE_BADDUALVIEW As Long = -6

Public Const EDS_RAWMODE As Long = &H2

Public Const EWX_LOGOFF As Long = 0
Public Const EWX_SHUTDOWN As Long = &H1
Public Const EWX_REBOOT As Long = &H2
Public Const EWX_FORCE As Long = &H4
Public Const EWX_POWEROFF As Long = &H8
Public Const EWX_FORCEIFHUNG As Long = &H10

Public Const FE_FONTSMOOTHINGSTANDARD As Long = &H1
Public Const FE_FONTSMOOTHINGCLEARTYPE As Long = &H2
Public Const FE_FONTSMOOTHINGDOCKING  As Long = &H8000

Public Const FKF_FILTERKEYSON As Long = &H1
Public Const FKF_AVAILABLE As Long = &H2
Public Const FKF_HOTKEYACTIVE As Long = &H4
Public Const FKF_CONFIRMHOTKEY As Long = &H8
Public Const FKF_HOTKEYSOUND As Long = &H10
Public Const FKF_INDICATOR As Long = &H20
Public Const FKF_CLICKON As Long = &H40

Public Const GA_PARENT As Long = 1
Public Const GA_ROOT As Long = 2
Public Const GA_ROOTOWNER As Long = 3

Public Const GR_GDIOBJECTS As Long = 0
Public Const GR_USEROBJECTS As Long = 1

Public Const GW_HWNDFIRST As Long = 0
Public Const GW_HWNDLAST As Long = 1
Public Const GW_HWNDNEXT As Long = 2
Public Const GW_HWNDPREV As Long = 3
Public Const GW_OWNER As Long = 4
Public Const GW_CHILD As Long = 5
Public Const GW_ENABLEDPOPUP As Long = 6
Public Const GW_MAX As Long = 6

Public Const GWL_WNDPROC As Long = -4
Public Const GWL_HINSTANCE As Long = -6
Public Const GWL_HWNDPARENT As Long = -8
Public Const GWL_STYLE As Long = -16
Public Const GWL_EXSTYLE As Long = -20
Public Const GWL_USERDATA As Long = -21
Public Const GWL_ID As Long = -12

Public Const HCF_HIGHCONTRASTON As Long = &H1
Public Const HCF_AVAILABLE As Long = &H2
Public Const HCF_HOTKEYACTIVE As Long = &H4
Public Const HCF_CONFIRMHOTKEY As Long = &H8
Public Const HCF_HOTKEYSOUND As Long = &H10
Public Const HCF_INDICATOR As Long = &H20
Public Const HCF_HOTKEYAVAILABLE As Long = &H40

Public Const HWND_TOP As Long = 0
Public Const HWND_BOTTOM As Long = 1
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2

Public Const KL_NAMELENGTH As Long = 9

Public Const MB_OK As Long = &H0
Public Const MB_OKCANCEL As Long = &H1
Public Const MB_ABORTRETRYIGNORE As Long = &H2
Public Const MB_YESNOCANCEL As Long = &H3
Public Const MB_YESNO As Long = &H4
Public Const MB_RETRYCANCEL As Long = &H5
Public Const MB_CANCELTRYCONTINUE As Long = &H6
Public Const MB_ICONHAND As Long = &H10
Public Const MB_ICONQUESTION As Long = &H20
Public Const MB_ICONEXCLAMATION As Long = &H30
Public Const MB_ICONASTERISK As Long = &H40
Public Const MB_USERICON As Long = &H80
Public Const MB_ICONWARNING As Long = MB_ICONEXCLAMATION
Public Const MB_ICONERROR As Long = MB_ICONHAND
Public Const MB_ICONINFORMATION As Long = MB_ICONASTERISK
Public Const MB_ICONSTOP As Long = MB_ICONHAND
Public Const MB_DEFBUTTON1 As Long = &H0
Public Const MB_DEFBUTTON2 As Long = &H100
Public Const MB_DEFBUTTON3 As Long = &H200
Public Const MB_DEFBUTTON4 As Long = &H300
Public Const MB_APPLMODAL As Long = &H0
Public Const MB_SYSTEMMODAL As Long = &H1000
Public Const MB_TASKMODAL As Long = &H2000
Public Const MB_HELP As Long = &H4000
Public Const MB_NOFOCUS As Long = &H8000
Public Const MB_SETFOREGROUND As Long = &H10000
Public Const MB_DEFAULT_DESKTOP_ONLY As Long = &H20000
Public Const MB_TOPMOST As Long = &H40000
Public Const MB_RIGHT As Long = &H80000
Public Const MB_RTLREADING As Long = &H100000
Public Const MB_SERVICE_NOTIFICATION As Long = &H200000
Public Const MB_SERVICE_NOTIFICATION_NT3X As Long = &H40000
Public Const MB_TYPEMASK As Long = &HF
Public Const MB_ICONMASK As Long = &HF0
Public Const MB_DEFMASK As Long = &HF00
Public Const MB_MODEMASK As Long = &H3000
Public Const MB_MISCMASK As Long = &HC000

Public Const MKF_MOUSEKEYSON As Long = &H1
Public Const MKF_AVAILABLE As Long = &H2
Public Const MKF_HOTKEYACTIVE As Long = &H4
Public Const MKF_CONFIRMHOTKEY As Long = &H8
Public Const MKF_HOTKEYSOUND As Long = &H10
Public Const MKF_INDICATOR As Long = &H20
Public Const MKF_MODIFIERS As Long = &H40
Public Const MKF_REPLACENUMBERS As Long = &H80
Public Const MKF_LEFTBUTTONSEL As Long = &H10000000
Public Const MKF_RIGHTBUTTONSEL As Long = &H20000000
Public Const MKF_LEFTBUTTONDOWN As Long = &H1000000
Public Const MKF_RIGHTBUTTONDOWN As Long = &H2000000
Public Const MKF_MOUSEMODE As Long = &H80000000

Public Const MONITORINFOF_PRIMARY As Long = &H1

Public Const SERKF_SERIALKEYSON As Long = &H1
Public Const SERKF_AVAILABLE As Long = &H2
Public Const SERKF_INDICATOR As Long = &H4

Public Const SKF_STICKYKEYSON As Long = &H1
Public Const SKF_AVAILABLE As Long = &H2
Public Const SKF_HOTKEYACTIVE As Long = &H4
Public Const SKF_CONFIRMHOTKEY As Long = &H8
Public Const SKF_HOTKEYSOUND As Long = &H10
Public Const SKF_INDICATOR As Long = &H20
Public Const SKF_AUDIBLEFEEDBACK As Long = &H40
Public Const SKF_TRISTATE As Long = &H80
Public Const SKF_TWOKEYSOFF As Long = &H100
Public Const SKF_LALTLATCHED As Long = &H10000000
Public Const SKF_LCTLLATCHED As Long = &H4000000
Public Const SKF_LSHIFTLATCHED As Long = &H1000000
Public Const SKF_RALTLATCHED As Long = &H20000000
Public Const SKF_RCTLLATCHED As Long = &H8000000
Public Const SKF_RSHIFTLATCHED As Long = &H2000000
Public Const SKF_LWINLATCHED As Long = &H40000000
Public Const SKF_RWINLATCHED As Long = &H80000000
Public Const SKF_LALTLOCKED As Long = &H100000
Public Const SKF_LCTLLOCKED As Long = &H40000
Public Const SKF_LSHIFTLOCKED As Long = &H10000
Public Const SKF_RALTLOCKED As Long = &H200000
Public Const SKF_RCTLLOCKED As Long = &H80000
Public Const SKF_RSHIFTLOCKED As Long = &H20000
Public Const SKF_LWINLOCKED As Long = &H400000
Public Const SKF_RWINLOCKED As Long = &H800000

Public Const SM_CXSCREEN As Long = 0
Public Const SM_CYSCREEN As Long = 1
Public Const SM_CXVSCROLL As Long = 2
Public Const SM_CYHSCROLL As Long = 3
Public Const SM_CYCAPTION As Long = 4
Public Const SM_CXBORDER As Long = 5
Public Const SM_CYBORDER As Long = 6
Public Const SM_CXDLGFRAME As Long = 7
Public Const SM_CYDLGFRAME As Long = 8
Public Const SM_CYVTHUMB As Long = 9
Public Const SM_CXHTHUMB As Long = 10
Public Const SM_CXICON As Long = 11
Public Const SM_CYICON As Long = 12
Public Const SM_CXCURSOR As Long = 13
Public Const SM_CYCURSOR As Long = 14
Public Const SM_CYMENU As Long = 15
Public Const SM_CXFULLSCREEN As Long = 16
Public Const SM_CYFULLSCREEN As Long = 17
Public Const SM_CYKANJIWINDOW As Long = 18
Public Const SM_MOUSEPRESENT As Long = 19
Public Const SM_CYVSCROLL As Long = 20
Public Const SM_CXHSCROLL As Long = 21
Public Const SM_DEBUG As Long = 22
Public Const SM_SWAPBUTTON As Long = 23
Public Const SM_RESERVED1 As Long = 24
Public Const SM_RESERVED2 As Long = 25
Public Const SM_RESERVED3 As Long = 26
Public Const SM_RESERVED4 As Long = 27
Public Const SM_CXMIN As Long = 28
Public Const SM_CYMIN As Long = 29
Public Const SM_CXSIZE As Long = 30
Public Const SM_CYSIZE As Long = 31
Public Const SM_CXFRAME As Long = 32
Public Const SM_CYFRAME As Long = 33
Public Const SM_CXMINTRACK As Long = 34
Public Const SM_CYMINTRACK As Long = 35
Public Const SM_CXDOUBLECLK As Long = 36
Public Const SM_CYDOUBLECLK As Long = 37
Public Const SM_CXICONSPACING As Long = 38
Public Const SM_CYICONSPACING As Long = 39
Public Const SM_MENUDROPALIGNMENT As Long = 40
Public Const SM_PENWINDOWS As Long = 41
Public Const SM_DBCSENABLED As Long = 42
Public Const SM_CMOUSEBUTTONS As Long = 43
Public Const SM_CXFIXEDFRAME As Long = SM_CXDLGFRAME
Public Const SM_CYFIXEDFRAME As Long = SM_CYDLGFRAME
Public Const SM_CXSIZEFRAME As Long = SM_CXFRAME
Public Const SM_CYSIZEFRAME As Long = SM_CYFRAME
Public Const SM_SECURE As Long = 44
Public Const SM_CXEDGE As Long = 45
Public Const SM_CYEDGE As Long = 46
Public Const SM_CXMINSPACING As Long = 47
Public Const SM_CYMINSPACING As Long = 48
Public Const SM_CXSMICON As Long = 49
Public Const SM_CYSMICON As Long = 50
Public Const SM_CYSMCAPTION As Long = 51
Public Const SM_CXSMSIZE As Long = 52
Public Const SM_CYSMSIZE As Long = 53
Public Const SM_CXMENUSIZE As Long = 54
Public Const SM_CYMENUSIZE As Long = 55
Public Const SM_ARRANGE As Long = 56
Public Const SM_CXMINIMIZED As Long = 57
Public Const SM_CYMINIMIZED As Long = 58
Public Const SM_CXMAXTRACK As Long = 59
Public Const SM_CYMAXTRACK As Long = 60
Public Const SM_CXMAXIMIZED As Long = 61
Public Const SM_CYMAXIMIZED As Long = 62
Public Const SM_NETWORK As Long = 63
Public Const SM_CLEANBOOT As Long = 67
Public Const SM_CXDRAG As Long = 68
Public Const SM_CYDRAG As Long = 69
Public Const SM_SHOWSOUNDS As Long = 70
Public Const SM_CXMENUCHECK As Long = 71
Public Const SM_CYMENUCHECK As Long = 72
Public Const SM_SLOWMACHINE As Long = 73
Public Const SM_MIDEASTENABLED As Long = 74
Public Const SM_MOUSEWHEELPRESENT As Long = 75
Public Const SM_XVIRTUALSCREEN As Long = 76
Public Const SM_YVIRTUALSCREEN As Long = 77
Public Const SM_CXVIRTUALSCREEN As Long = 78
Public Const SM_CYVIRTUALSCREEN As Long = 79
Public Const SM_CMONITORS As Long = 80
Public Const SM_SAMEDISPLAYFORMAT As Long = 81
Public Const SM_IMMENABLED As Long = 82
Public Const SM_CXFOCUSBORDER As Long = 83
Public Const SM_CYFOCUSBORDER As Long = 84
Public Const SM_CMETRICS As Long = 86
Public Const SM_REMOTESESSION As Long = &H1000
Public Const SM_SHUTTINGDOWN As Long = &H2000

Public Const SPI_GETBEEP As Long = 1
Public Const SPI_SETBEEP As Long = 2
Public Const SPI_GETMOUSE As Long = 3
Public Const SPI_SETMOUSE As Long = 4
Public Const SPI_GETBORDER As Long = 5
Public Const SPI_SETBORDER As Long = 6
Public Const SPI_GETKEYBOARDSPEED As Long = 10
Public Const SPI_SETKEYBOARDSPEED As Long = 11
Public Const SPI_LANGDRIVER As Long = 12
Public Const SPI_ICONHORIZONTALSPACING As Long = 13
Public Const SPI_GETSCREENSAVETIMEOUT As Long = 14
Public Const SPI_SETSCREENSAVETIMEOUT As Long = 15
Public Const SPI_GETSCREENSAVEACTIVE As Long = 16
Public Const SPI_SETSCREENSAVEACTIVE As Long = 17
Public Const SPI_GETGRIDGRANULARITY As Long = 18
Public Const SPI_SETGRIDGRANULARITY As Long = 19
Public Const SPI_SETDESKWALLPAPER As Long = 20
Public Const SPI_SETDESKPATTERN As Long = 21
Public Const SPI_GETKEYBOARDDELAY As Long = 22
Public Const SPI_SETKEYBOARDDELAY As Long = 23
Public Const SPI_ICONVERTICALSPACING As Long = 24
Public Const SPI_GETICONTITLEWRAP As Long = 25
Public Const SPI_SETICONTITLEWRAP As Long = 26
Public Const SPI_GETMENUDROPALIGNMENT As Long = 27
Public Const SPI_SETMENUDROPALIGNMENT As Long = 28
Public Const SPI_SETDOUBLECLKWIDTH As Long = 29
Public Const SPI_SETDOUBLECLKHEIGHT As Long = 30
Public Const SPI_GETICONTITLELOGFONT As Long = 31
Public Const SPI_SETDOUBLECLICKTIME As Long = 32
Public Const SPI_SETMOUSEBUTTONSWAP As Long = 33
Public Const SPI_SETICONTITLELOGFONT As Long = 34
Public Const SPI_GETFASTTASKSWITCH As Long = 35
Public Const SPI_SETFASTTASKSWITCH As Long = 36
Public Const SPI_SETDRAGFULLWINDOWS As Long = 37
Public Const SPI_GETDRAGFULLWINDOWS As Long = 38
Public Const SPI_GETNONCLIENTMETRICS As Long = 41
Public Const SPI_SETNONCLIENTMETRICS As Long = 42
Public Const SPI_GETMINIMIZEDMETRICS As Long = 43
Public Const SPI_SETMINIMIZEDMETRICS As Long = 44
Public Const SPI_GETICONMETRICS As Long = 45
Public Const SPI_SETICONMETRICS As Long = 46
Public Const SPI_SETWORKAREA As Long = 47
Public Const SPI_GETWORKAREA As Long = 48
Public Const SPI_SETPENWINDOWS As Long = 49
Public Const SPI_GETHIGHCONTRAST As Long = 66
Public Const SPI_SETHIGHCONTRAST As Long = 67
Public Const SPI_GETKEYBOARDPREF As Long = 68
Public Const SPI_SETKEYBOARDPREF As Long = 69
Public Const SPI_GETSCREENREADER As Long = 70
Public Const SPI_SETSCREENREADER As Long = 71
Public Const SPI_GETANIMATION As Long = 72
Public Const SPI_SETANIMATION As Long = 73
Public Const SPI_GETFONTSMOOTHING As Long = 74
Public Const SPI_SETFONTSMOOTHING As Long = 75
Public Const SPI_SETDRAGWIDTH As Long = 76
Public Const SPI_SETDRAGHEIGHT As Long = 77
Public Const SPI_SETHANDHELD As Long = 78
Public Const SPI_GETLOWPOWERTIMEOUT As Long = 79
Public Const SPI_GETPOWEROFFTIMEOUT As Long = 80
Public Const SPI_SETLOWPOWERTIMEOUT As Long = 81
Public Const SPI_SETPOWEROFFTIMEOUT As Long = 82
Public Const SPI_GETLOWPOWERACTIVE As Long = 83
Public Const SPI_GETPOWEROFFACTIVE As Long = 84
Public Const SPI_SETLOWPOWERACTIVE As Long = 85
Public Const SPI_SETPOWEROFFACTIVE As Long = 86
Public Const SPI_SETCURSORS As Long = 87
Public Const SPI_SETICONS As Long = 88
Public Const SPI_GETDEFAULTINPUTLANG As Long = 89
Public Const SPI_SETDEFAULTINPUTLANG As Long = 90
Public Const SPI_SETLANGTOGGLE As Long = 91
Public Const SPI_GETWINDOWSEXTENSION As Long = 92
Public Const SPI_SETMOUSETRAILS As Long = 93
Public Const SPI_GETMOUSETRAILS As Long = 94
Public Const SPI_SETSCREENSAVERRUNNING As Long = 97
Public Const SPI_SCREENSAVERRUNNING As Long = SPI_SETSCREENSAVERRUNNING
Public Const SPI_GETFILTERKEYS As Long = 50
Public Const SPI_SETFILTERKEYS As Long = 51
Public Const SPI_GETTOGGLEKEYS As Long = 52
Public Const SPI_SETTOGGLEKEYS As Long = 53
Public Const SPI_GETMOUSEKEYS As Long = 54
Public Const SPI_SETMOUSEKEYS As Long = 55
Public Const SPI_GETSHOWSOUNDS As Long = 56
Public Const SPI_SETSHOWSOUNDS As Long = 57
Public Const SPI_GETSTICKYKEYS As Long = 58
Public Const SPI_SETSTICKYKEYS As Long = 59
Public Const SPI_GETACCESSTIMEOUT As Long = 60
Public Const SPI_SETACCESSTIMEOUT As Long = 61
Public Const SPI_GETSERIALKEYS As Long = 62
Public Const SPI_SETSERIALKEYS As Long = 63
Public Const SPI_GETSOUNDSENTRY As Long = 64
Public Const SPI_SETSOUNDSENTRY As Long = 65
Public Const SPI_GETSNAPTODEFBUTTON As Long = 95
Public Const SPI_SETSNAPTODEFBUTTON As Long = 96
Public Const SPI_GETMOUSEHOVERWIDTH As Long = 98
Public Const SPI_SETMOUSEHOVERWIDTH As Long = 99
Public Const SPI_GETMOUSEHOVERHEIGHT As Long = 100
Public Const SPI_SETMOUSEHOVERHEIGHT As Long = 101
Public Const SPI_GETMOUSEHOVERTIME As Long = 102
Public Const SPI_SETMOUSEHOVERTIME As Long = 103
Public Const SPI_GETWHEELSCROLLLINES As Long = 104
Public Const SPI_SETWHEELSCROLLLINES As Long = 105
Public Const SPI_GETMENUSHOWDELAY As Long = 106
Public Const SPI_SETMENUSHOWDELAY As Long = 107
Public Const SPI_GETSHOWIMEUI As Long = 110
Public Const SPI_SETSHOWIMEUI As Long = 111
Public Const SPI_GETMOUSESPEED As Long = 112
Public Const SPI_SETMOUSESPEED As Long = 113
Public Const SPI_GETSCREENSAVERRUNNING As Long = 114
Public Const SPI_GETDESKWALLPAPER As Long = 115
Public Const SPI_GETACTIVEWINDOWTRACKING As Long = &H1000
Public Const SPI_SETACTIVEWINDOWTRACKING As Long = &H1001
Public Const SPI_GETMENUANIMATION As Long = &H1002
Public Const SPI_SETMENUANIMATION As Long = &H1003
Public Const SPI_GETCOMBOBOXANIMATION As Long = &H1004
Public Const SPI_SETCOMBOBOXANIMATION As Long = &H1005
Public Const SPI_GETLISTBOXSMOOTHSCROLLING As Long = &H1006
Public Const SPI_SETLISTBOXSMOOTHSCROLLING As Long = &H1007
Public Const SPI_GETGRADIENTCAPTIONS As Long = &H1008
Public Const SPI_SETGRADIENTCAPTIONS As Long = &H1009
Public Const SPI_GETKEYBOARDCUES As Long = &H100A
Public Const SPI_SETKEYBOARDCUES As Long = &H100B
Public Const SPI_GETMENUUNDERLINES As Long = SPI_GETKEYBOARDCUES
Public Const SPI_SETMENUUNDERLINES As Long = SPI_SETKEYBOARDCUES
Public Const SPI_GETACTIVEWNDTRKZORDER As Long = &H100C
Public Const SPI_SETACTIVEWNDTRKZORDER As Long = &H100D
Public Const SPI_GETHOTTRACKING As Long = &H100E
Public Const SPI_SETHOTTRACKING As Long = &H100F
Public Const SPI_GETMENUFADE As Long = &H1012
Public Const SPI_SETMENUFADE As Long = &H1013
Public Const SPI_GETSELECTIONFADE As Long = &H1014
Public Const SPI_SETSELECTIONFADE As Long = &H1015
Public Const SPI_GETTOOLTIPANIMATION As Long = &H1016
Public Const SPI_SETTOOLTIPANIMATION As Long = &H1017
Public Const SPI_GETTOOLTIPFADE As Long = &H1018
Public Const SPI_SETTOOLTIPFADE As Long = &H1019
Public Const SPI_GETCURSORSHADOW As Long = &H101A
Public Const SPI_SETCURSORSHADOW As Long = &H101B
Public Const SPI_GETMOUSESONAR As Long = &H101C
Public Const SPI_SETMOUSESONAR As Long = &H101D
Public Const SPI_GETMOUSECLICKLOCK As Long = &H101E
Public Const SPI_SETMOUSECLICKLOCK As Long = &H101F
Public Const SPI_GETMOUSEVANISH As Long = &H1020
Public Const SPI_SETMOUSEVANISH As Long = &H1021
Public Const SPI_GETFLATMENU As Long = &H1022
Public Const SPI_SETFLATMENU As Long = &H1023
Public Const SPI_GETDROPSHADOW As Long = &H1024
Public Const SPI_SETDROPSHADOW As Long = &H1025
Public Const SPI_GETUIEFFECTS As Long = &H103E
Public Const SPI_SETUIEFFECTS As Long = &H103F
Public Const SPI_GETFOREGROUNDLOCKTIMEOUT As Long = &H2000
Public Const SPI_SETFOREGROUNDLOCKTIMEOUT As Long = &H2001
Public Const SPI_GETACTIVEWNDTRKTIMEOUT As Long = &H2002
Public Const SPI_SETACTIVEWNDTRKTIMEOUT As Long = &H2003
Public Const SPI_GETFOREGROUNDFLASHCOUNT As Long = &H2004
Public Const SPI_SETFOREGROUNDFLASHCOUNT As Long = &H2005
Public Const SPI_GETCARETWIDTH As Long = &H2006
Public Const SPI_SETCARETWIDTH As Long = &H2007
Public Const SPI_GETMOUSECLICKLOCKTIME As Long = &H2008
Public Const SPI_SETMOUSECLICKLOCKTIME As Long = &H2009
Public Const SPI_GETFONTSMOOTHINGTYPE As Long = &H200A
Public Const SPI_SETFONTSMOOTHINGTYPE As Long = &H200B
Public Const SPI_GETFONTSMOOTHINGCONTRAST As Long = &H200C
Public Const SPI_SETFONTSMOOTHINGCONTRAST As Long = &H200D
Public Const SPI_GETFOCUSBORDERWIDTH As Long = &H200E
Public Const SPI_SETFOCUSBORDERWIDTH As Long = &H200F
Public Const SPI_GETFOCUSBORDERHEIGHT As Long = &H2010
Public Const SPI_SETFOCUSBORDERHEIGHT As Long = &H2011
Public Const SPIF_UPDATEINIFILE As Long = &H1
Public Const SPIF_SENDWININICHANGE As Long = &H2
Public Const SPIF_SENDCHANGE As Long = SPIF_SENDWININICHANGE

Public Const SSF_SOUNDSENTRYON As Long = &H1
Public Const SSF_AVAILABLE As Long = &H2
Public Const SSF_INDICATOR As Long = &H4

Public Const SSGF_NONE As Long = 0
Public Const SSGF_DISPLAY As Long = 3

Public Const SSTF_NONE As Long = 0
Public Const SSTF_CHARS As Long = 1
Public Const SSTF_BORDER As Long = 2
Public Const SSTF_DISPLAY As Long = 3

Public Const SSWF_NONE As Long = 0
Public Const SSWF_TITLE As Long = 1
Public Const SSWF_WINDOW As Long = 2
Public Const SSWF_DISPLAY As Long = 3
Public Const SSWF_CUSTOM As Long = 4

Public Const SW_HIDE As Long = 0
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_NORMAL As Long = 1
Public Const SW_SHOWMINIMIZED As Long = 2
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_MAXIMIZE As Long = 3
Public Const SW_SHOWNOACTIVATE As Long = 4
Public Const SW_SHOW As Long = 5
Public Const SW_MINIMIZE As Long = 6
Public Const SW_SHOWMINNOACTIVE As Long = 7
Public Const SW_SHOWNA As Long = 8
Public Const SW_RESTORE As Long = 9
Public Const SW_SHOWDEFAULT As Long = 10
Public Const SW_FORCEMINIMIZE As Long = 11
Public Const SW_MAX As Long = 11

Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_NOREDRAW As Long = &H8
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_FRAMECHANGED As Long = &H20
Public Const SWP_SHOWWINDOW As Long = &H40
Public Const SWP_HIDEWINDOW As Long = &H80
Public Const SWP_NOCOPYBITS As Long = &H100
Public Const SWP_NOOWNERZORDER As Long = &H200
Public Const SWP_NOSENDCHANGING As Long = &H400
Public Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION As Long = SWP_NOOWNERZORDER
Public Const SWP_DEFERERASE As Long = &H2000
Public Const SWP_ASYNCWINDOWPOS As Long = &H4000

Public Const TKF_TOGGLEKEYSON As Long = &H1
Public Const TKF_AVAILABLE As Long = &H2
Public Const TKF_HOTKEYACTIVE As Long = &H4
Public Const TKF_CONFIRMHOTKEY As Long = &H8
Public Const TKF_HOTKEYSOUND As Long = &H10
Public Const TKF_INDICATOR As Long = &H20

Public Const WM_NULL As Long = &H0
Public Const WM_CREATE As Long = &H1
Public Const WM_DESTROY As Long = &H2
Public Const WM_MOVE As Long = &H3
Public Const WM_SIZE As Long = &H5
Public Const WM_ACTIVATE As Long = &H6
Public Const WM_SETFOCUS As Long = &H7
Public Const WM_KILLFOCUS As Long = &H8
Public Const WM_ENABLE As Long = &HA
Public Const WM_SETREDRAW As Long = &HB
Public Const WM_SETTEXT As Long = &HC
Public Const WM_GETTEXT As Long = &HD
Public Const WM_GETTEXTLENGTH As Long = &HE
Public Const WM_PAINT As Long = &HF
Public Const WM_CLOSE As Long = &H10
Public Const WM_QUERYENDSESSION As Long = &H11
Public Const WM_QUERYOPEN As Long = &H13
Public Const WM_ENDSESSION As Long = &H16
Public Const WM_QUIT As Long = &H12
Public Const WM_ERASEBKGND As Long = &H14
Public Const WM_SYSCOLORCHANGE As Long = &H15
Public Const WM_SHOWWINDOW As Long = &H18
Public Const WM_WININICHANGE As Long = &H1A
Public Const WM_SETTINGCHANGE As Long = WM_WININICHANGE
Public Const WM_DEVMODECHANGE As Long = &H1B
Public Const WM_ACTIVATEAPP As Long = &H1C
Public Const WM_FONTCHANGE As Long = &H1D
Public Const WM_TIMECHANGE As Long = &H1E
Public Const WM_CANCELMODE As Long = &H1F
Public Const WM_SETCURSOR As Long = &H20
Public Const WM_MOUSEACTIVATE As Long = &H21
Public Const WM_CHILDACTIVATE As Long = &H22
Public Const WM_QUEUESYNC As Long = &H23
Public Const WM_GETMINMAXINFO As Long = &H24
Public Const WM_PAINTICON As Long = &H26
Public Const WM_ICONERASEBKGND As Long = &H27
Public Const WM_NEXTDLGCTL As Long = &H28
Public Const WM_SPOOLERSTATUS As Long = &H2A
Public Const WM_DRAWITEM As Long = &H2B
Public Const WM_MEASUREITEM As Long = &H2C
Public Const WM_DELETEITEM As Long = &H2D
Public Const WM_VKEYTOITEM As Long = &H2E
Public Const WM_CHARTOITEM As Long = &H2F
Public Const WM_SETFONT As Long = &H30
Public Const WM_GETFONT As Long = &H31
Public Const WM_SETHOTKEY As Long = &H32
Public Const WM_GETHOTKEY As Long = &H33
Public Const WM_QUERYDRAGICON As Long = &H37
Public Const WM_COMPAREITEM As Long = &H39
Public Const WM_GETOBJECT As Long = &H3D
Public Const WM_COMPACTING As Long = &H41
Public Const WM_COMMNOTIFY As Long = &H44
Public Const WM_WINDOWPOSCHANGING As Long = &H46
Public Const WM_WINDOWPOSCHANGED As Long = &H47
Public Const WM_POWER As Long = &H48
Public Const WM_COPYDATA As Long = &H4A
Public Const WM_CANCELJOURNAL As Long = &H4B
Public Const WM_NOTIFY As Long = &H4E
Public Const WM_INPUTLANGCHANGEREQUEST As Long = &H50
Public Const WM_INPUTLANGCHANGE As Long = &H51
Public Const WM_TCARD As Long = &H52
Public Const WM_HELP As Long = &H53
Public Const WM_USERCHANGED As Long = &H54
Public Const WM_NOTIFYFORMAT As Long = &H55
Public Const WM_CONTEXTMENU As Long = &H7B
Public Const WM_STYLECHANGING As Long = &H7C
Public Const WM_STYLECHANGED As Long = &H7D
Public Const WM_DISPLAYCHANGE As Long = &H7E
Public Const WM_GETICON As Long = &H7F
Public Const WM_SETICON As Long = &H80
Public Const WM_NCCREATE As Long = &H81
Public Const WM_NCDESTROY As Long = &H82
Public Const WM_NCCALCSIZE As Long = &H83
Public Const WM_NCHITTEST As Long = &H84
Public Const WM_NCPAINT As Long = &H85
Public Const WM_NCACTIVATE As Long = &H86
Public Const WM_GETDLGCODE As Long = &H87
Public Const WM_SYNCPAINT As Long = &H88
Public Const WM_NCMOUSEMOVE As Long = &HA0
Public Const WM_NCLBUTTONDOWN As Long = &HA1
Public Const WM_NCLBUTTONUP As Long = &HA2
Public Const WM_NCLBUTTONDBLCLK As Long = &HA3
Public Const WM_NCRBUTTONDOWN As Long = &HA4
Public Const WM_NCRBUTTONUP As Long = &HA5
Public Const WM_NCRBUTTONDBLCLK As Long = &HA6
Public Const WM_NCMBUTTONDOWN As Long = &HA7
Public Const WM_NCMBUTTONUP As Long = &HA8
Public Const WM_NCMBUTTONDBLCLK As Long = &HA9
Public Const WM_NCXBUTTONDOWN As Long = &HAB
Public Const WM_NCXBUTTONUP As Long = &HAC
Public Const WM_NCXBUTTONDBLCLK As Long = &HAD
Public Const WM_INPUT As Long = &HFF
Public Const WM_KEYFIRST As Long = &H100
Public Const WM_KEYDOWN As Long = &H100
Public Const WM_KEYUP As Long = &H101
Public Const WM_CHAR As Long = &H102
Public Const WM_DEADCHAR As Long = &H103
Public Const WM_SYSKEYDOWN As Long = &H104
Public Const WM_SYSKEYUP As Long = &H105
Public Const WM_SYSCHAR As Long = &H106
Public Const WM_SYSDEADCHAR As Long = &H107
Public Const WM_KEYLAST As Long = &H108
Public Const WM_IME_STARTCOMPOSITION As Long = &H10D
Public Const WM_IME_ENDCOMPOSITION As Long = &H10E
Public Const WM_IME_COMPOSITION As Long = &H10F
Public Const WM_IME_KEYLAST As Long = &H10F
Public Const WM_INITDIALOG As Long = &H110
Public Const WM_COMMAND As Long = &H111
Public Const WM_SYSCOMMAND As Long = &H112
Public Const WM_TIMER As Long = &H113
Public Const WM_HSCROLL As Long = &H114
Public Const WM_VSCROLL As Long = &H115
Public Const WM_INITMENU As Long = &H116
Public Const WM_INITMENUPOPUP As Long = &H117
Public Const WM_MENUSELECT As Long = &H11F
Public Const WM_MENUCHAR As Long = &H120
Public Const WM_ENTERIDLE As Long = &H121
Public Const WM_MENURBUTTONUP As Long = &H122
Public Const WM_MENUDRAG As Long = &H123
Public Const WM_MENUGETOBJECT As Long = &H124
Public Const WM_UNINITMENUPOPUP As Long = &H125
Public Const WM_MENUCOMMAND As Long = &H126
Public Const WM_CHANGEUISTATE As Long = &H127
Public Const WM_UPDATEUISTATE As Long = &H128
Public Const WM_QUERYUISTATE As Long = &H129
Public Const WM_CTLCOLORMSGBOX As Long = &H132
Public Const WM_CTLCOLOREDIT As Long = &H133
Public Const WM_CTLCOLORLISTBOX As Long = &H134
Public Const WM_CTLCOLORBTN As Long = &H135
Public Const WM_CTLCOLORDLG As Long = &H136
Public Const WM_CTLCOLORSCROLLBAR As Long = &H137
Public Const WM_CTLCOLORSTATIC As Long = &H138
Public Const WM_MOUSEFIRST As Long = &H200
Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209
Public Const WM_MOUSEWHEEL As Long = &H20A
Public Const WM_XBUTTONDOWN As Long = &H20B
Public Const WM_XBUTTONUP As Long = &H20C
Public Const WM_XBUTTONDBLCLK As Long = &H20D
Public Const WM_MOUSELAST As Long = &H20D
Public Const WM_PARENTNOTIFY As Long = &H210
Public Const WM_ENTERMENULOOP As Long = &H211
Public Const WM_EXITMENULOOP As Long = &H212
Public Const WM_NEXTMENU As Long = &H213
Public Const WM_SIZING As Long = &H214
Public Const WM_CAPTURECHANGED As Long = &H215
Public Const WM_MOVING As Long = &H216
Public Const WM_POWERBROADCAST As Long = &H218
Public Const WM_DEVICECHANGE As Long = &H219
Public Const WM_MDICREATE As Long = &H220
Public Const WM_MDIDESTROY As Long = &H221
Public Const WM_MDIACTIVATE As Long = &H222
Public Const WM_MDIRESTORE As Long = &H223
Public Const WM_MDINEXT As Long = &H224
Public Const WM_MDIMAXIMIZE As Long = &H225
Public Const WM_MDITILE As Long = &H226
Public Const WM_MDICASCADE As Long = &H227
Public Const WM_MDIICONARRANGE As Long = &H228
Public Const WM_MDIGETACTIVE As Long = &H229
Public Const WM_MDISETMENU As Long = &H230
Public Const WM_ENTERSIZEMOVE As Long = &H231
Public Const WM_EXITSIZEMOVE As Long = &H232
Public Const WM_DROPFILES As Long = &H233
Public Const WM_MDIREFRESHMENU As Long = &H234
Public Const WM_IME_SETCONTEXT As Long = &H281
Public Const WM_IME_NOTIFY As Long = &H282
Public Const WM_IME_CONTROL As Long = &H283
Public Const WM_IME_COMPOSITIONFULL As Long = &H284
Public Const WM_IME_SELECT As Long = &H285
Public Const WM_IME_CHAR As Long = &H286
Public Const WM_IME_REQUEST As Long = &H288
Public Const WM_IME_KEYDOWN As Long = &H290
Public Const WM_IME_KEYUP As Long = &H291
Public Const WM_MOUSEHOVER As Long = &H2A1
Public Const WM_MOUSELEAVE As Long = &H2A3
Public Const WM_NCMOUSEHOVER As Long = &H2A0
Public Const WM_NCMOUSELEAVE As Long = &H2A2
Public Const WM_WTSSESSION_CHANGE As Long = &H2B1
Public Const WM_TABLET_FIRST As Long = &H2C0
Public Const WM_TABLET_LAST As Long = &H2DF
Public Const WM_CUT As Long = &H300
Public Const WM_COPY As Long = &H301
Public Const WM_PASTE As Long = &H302
Public Const WM_CLEAR As Long = &H303
Public Const WM_UNDO As Long = &H304
Public Const WM_RENDERFORMAT As Long = &H305
Public Const WM_RENDERALLFORMATS As Long = &H306
Public Const WM_DESTROYCLIPBOARD As Long = &H307
Public Const WM_DRAWCLIPBOARD As Long = &H308
Public Const WM_PAINTCLIPBOARD As Long = &H309
Public Const WM_VSCROLLCLIPBOARD As Long = &H30A
Public Const WM_SIZECLIPBOARD As Long = &H30B
Public Const WM_ASKCBFORMATNAME As Long = &H30C
Public Const WM_CHANGECBCHAIN As Long = &H30D
Public Const WM_HSCROLLCLIPBOARD As Long = &H30E
Public Const WM_QUERYNEWPALETTE As Long = &H30F
Public Const WM_PALETTEISCHANGING As Long = &H310
Public Const WM_PALETTECHANGED As Long = &H311
Public Const WM_HOTKEY As Long = &H312
Public Const WM_PRINT As Long = &H317
Public Const WM_PRINTCLIENT As Long = &H318
Public Const WM_APPCOMMAND As Long = &H319
Public Const WM_THEMECHANGED As Long = &H31A
Public Const WM_HANDHELDFIRST As Long = &H358
Public Const WM_HANDHELDLAST As Long = &H35F
Public Const WM_AFXFIRST As Long = &H360
Public Const WM_AFXLAST As Long = &H37F
Public Const WM_PENWINFIRST As Long = &H380
Public Const WM_PENWINLAST As Long = &H38F
Public Const WM_APP As Long = &H8000
Public Const WM_USER As Long = &H400

Public Const WPF_SETMINPOSITION As Long = &H1
Public Const WPF_RESTORETOMAXIMIZED As Long = &H2
Public Const WPF_ASYNCWINDOWPLACEMENT As Long = &H4

Public Const WS_OVERLAPPED As Long = &H0
Public Const WS_POPUP As Long = &H80000000
Public Const WS_CHILD As Long = &H40000000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_DISABLED As Long = &H8000000
Public Const WS_CLIPSIBLINGS As Long = &H4000000
Public Const WS_CLIPCHILDREN As Long = &H2000000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_BORDER As Long = &H800000
Public Const WS_DLGFRAME As Long = &H400000
Public Const WS_VSCROLL As Long = &H200000
Public Const WS_HSCROLL As Long = &H100000
Public Const WS_SYSMENU As Long = &H80000
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_GROUP As Long = &H20000
Public Const WS_TABSTOP As Long = &H10000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_TILED As Long = WS_OVERLAPPED
Public Const WS_ICONIC As Long = WS_MINIMIZE
Public Const WS_SIZEBOX As Long = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUPWINDOW As Long = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW As Long = WS_CHILD
Public Const WS_TILEDWINDOW As Long = WS_OVERLAPPEDWINDOW

Public Const WS_EX_DLGMODALFRAME As Long = &H1
Public Const WS_EX_NOPARENTNOTIFY As Long = &H4
Public Const WS_EX_TOPMOST As Long = &H8
Public Const WS_EX_ACCEPTFILES As Long = &H10
Public Const WS_EX_TRANSPARENT As Long = &H20
Public Const WS_EX_MDICHILD As Long = &H40
Public Const WS_EX_TOOLWINDOW As Long = &H80
Public Const WS_EX_WINDOWEDGE As Long = &H100
Public Const WS_EX_CLIENTEDGE As Long = &H200
Public Const WS_EX_CONTEXTHELP As Long = &H400
Public Const WS_EX_RIGHT As Long = &H1000
Public Const WS_EX_LEFT As Long = &H0
Public Const WS_EX_RTLREADING As Long = &H2000
Public Const WS_EX_LTRREADING As Long = &H0
Public Const WS_EX_LEFTSCROLLBAR As Long = &H4000
Public Const WS_EX_RIGHTSCROLLBAR As Long = &H0
Public Const WS_EX_CONTROLPARENT As Long = &H10000
Public Const WS_EX_STATICEDGE As Long = &H20000
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_OVERLAPPEDWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_NOINHERITLAYOUT As Long = &H100000
Public Const WS_EX_LAYOUTRTL As Long = &H400000
Public Const WS_EX_COMPOSITED As Long = &H2000000
Public Const WS_EX_NOACTIVATE As Long = &H8000000

Public Const XBUTTON1 As Long = &H1
Public Const XBUTTON2 As Long = &H2


Public Type ACCESSTIMEOUT
    cbSize As Long
    dwFlags As Long
    iTimeOutMSec As Long
End Type

Public Type FILTERKEYS
    cbSize As Long
    dwFlags As Long
    iWaitMSec As Long
    iDelayMSec As Long
    iRepeatMSec As Long
    iBounceMSec As Long
End Type

Public Type HIGHCONTRAST
    cbSize As Long
    dwFlags As Long
    lpszDefaultScheme As String
End Type

Public Type ICONMETRICS
    cbSize As Long
    iHorzSpacing As Long
    iVertSpacing As Long
    iTitleWrap As Long
    lfFont As LOGFONT
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
    
Public Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As LARGE_INTEGER
    ullAvailPhys As LARGE_INTEGER
    ullTotalPageFile As LARGE_INTEGER
    ullAvailPageFile As LARGE_INTEGER
    ullTotalVirtual As LARGE_INTEGER
    ullAvailVirtual As LARGE_INTEGER
    ullAvailExtendedVirtual As LARGE_INTEGER
End Type

Public Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

Public Type MOUSEKEYS
    cbSize As Long
    dwFlags As Long
    iMaxSpeed As Long
    iTimeToMaxSpeed As Long
    iCtrlSpeed As Long
    dwReserved1 As Long
    dwReserved2 As Long
End Type

Public Type MONITORINFOEX
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
    szDevice As String * CCHDEVICENAME
End Type

Public Type SERIALKEYS
    cbSize As Long
    dwFlags As Long
    lpszActivePort As String
    lpszPort As String
    iBaudRate As Long
    iPortState As Long
    iActive As Long
End Type

Public Type SOUNDSENTRY
    cbSize As Long
    dwFlags As Long
    iFSTextEffect As Long
    iFSTextEffectMSec As Long
    iFSTextEffectColorBits As Long
    iFSGrafEffect As Long
    iFSGrafEffectMSec As Long
    iFSGrafEffectColor As Long
    iWindowsEffect As Long
    iWindowsEffectMSec As Long
    lpszWindowsEffectDLL As String
    iWindowsEffectOrdinal As Long
End Type

Public Type STICKYKEYS
    cbSize As Long
    dwFlags As Long
End Type

Public Type TOGGLEKEYS
    cbSize As Long
    dwFlags As Long
End Type

Public Type WINDOWINFO
    cbSize As Long
    rcWindow As RECT
    rcClient As RECT
    dwStyle As Long
    dwExStyle As Long
    dwWindowStatus As Long
    cxWindowBorders As Long
    cyWindowBorders As Long
    atomWindowType As Long
    wCreatorVersion As Integer
End Type

Public Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
