Attribute VB_Name = "mdlWindows"
Option Explicit


Const sLocation As String = "mdlWindows"


Public Function ComputerName_Get() As String
On Error GoTo VB_Error
    
    ComputerName_Get = String$(MAX_COMPUTERNAME_LENGTH + 1, 0)
    
    If GetComputerName(ComputerName_Get, MAX_COMPUTERNAME_LENGTH + 1) = False Then Call Error_API(Err.LastDllError, sLocation & "\ComputerName_Get", "GetComputerName")
    ComputerName_Get = Str_NullTerm_Fix(ComputerName_Get)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\ComputerName_Get")
Resume Next
End Function

Public Function Dir_Path_Get(ByVal lFolder As Long) As String
On Error GoTo VB_Error
    
    Dim sPath As String
    sPath = String$(MAX_PATH, 0)
    
    If Function_Exist("shell32.dll", "SHGetFolderPathA") = True Then
        If SHGetFolderPath(0&, lFolder, 0&, SHGFP_TYPE.SHGFP_TYPE_CURRENT, sPath) <> S_OK Then Call Error_API(Err.LastDllError, sLocation & "\Dir_Path_Get", "SHGetFolderPathA")
        Dir_Path_Get = Str_BckSlhTerm_Fix(Str_NullTerm_Fix(sPath))
    Else
        If Function_Exist("shell32.dll", "SHGetSpecialFolderPathA") = True Then
            If SHGetSpecialFolderPath(0&, sPath, lFolder, False) = False Then Call Error_API(Err.LastDllError, sLocation & "\Dir_Path_Get", "SHGetSpecialFolderPath")
            Dir_Path_Get = Str_BckSlhTerm_Fix(Str_NullTerm_Fix(sPath))
        End If
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Dir_Path_Get")
Resume Next
End Function

Public Function Function_Exist(ByVal sModule As String, ByVal sFunction As String) As Boolean
On Error GoTo VB_Error

    Dim hHandle As Long
    
    hHandle = GetModuleHandle(sModule)
    If hHandle = 0 Then
        If Err.LastDllError <> ERROR_MOD_NOT_FOUND Then Call Error_API(Err.LastDllError, sLocation & "\Function_Exist", "GetModuleHandle")
        
        hHandle = LoadLibraryEx(sModule, 0&, 0&): If hHandle = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Function_Exist", "LoadLibrary")
        
        If GetProcAddress(hHandle, sFunction) = 0 Then
            Call Error_API(Err.LastDllError, sLocation & "\Function_Exist", "GetProcAddress")
            Function_Exist = False
        Else
            Function_Exist = True
        End If
        
        If FreeLibrary(hHandle) = False Then Call Error_API(Err.LastDllError, sLocation & "\Function_Exist", "FreeLibrary")
    Else
        If GetProcAddress(hHandle, sFunction) = 0 Then
            Call Error_API(Err.LastDllError, sLocation & "\Function_Exist", "GetProcAddress")
            Function_Exist = Function_Exist
        Else
            Function_Exist = True
        End If
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Function_Exist")
Resume Next
End Function

Public Function LangIdent(ByVal lCode As Long) As String
On Error GoTo VB_Error
    
    Select Case lCode
        Case &H0: LangIdent = "Language Neutral"
        Case &H400: LangIdent = "Process Default Language"
        Case &H436: LangIdent = "Afrikaans"
        Case &H41C: LangIdent = "Albanian"
        Case &H401: LangIdent = "Arabic (Saudi Arabia)"
        Case &H801: LangIdent = "Arabic (Iraq)"
        Case &HC01: LangIdent = "Arabic (Egypt)"
        Case &H1001: LangIdent = "Arabic (Libya)"
        Case &H1401: LangIdent = "Arabic (Algeria)"
        Case &H1801: LangIdent = "Arabic (Morocco)"
        Case &H1C01: LangIdent = "Arabic (Tunisia)"
        Case &H2001: LangIdent = "Arabic (Oman)"
        Case &H2401: LangIdent = "Arabic (Yemen)"
        Case &H2801: LangIdent = "Arabic (Syria)"
        Case &H2C01: LangIdent = "Arabic (Jordan)"
        Case &H3001: LangIdent = "Arabic (Lebanon)"
        Case &H3401: LangIdent = "Arabic (Kuwait)"
        Case &H3801: LangIdent = "Arabic (U.A.E.)"
        Case &H3C01: LangIdent = "Arabic (Bahrain)"
        Case &H4001: LangIdent = "Arabic (Qatar)"
        Case &H42B: LangIdent = "Armenian"
        Case &H44D: LangIdent = "Assamese"
        Case &H42C: LangIdent = "Azeri (Latin)"
        Case &H82C: LangIdent = "Azeri (Cyrillic)"
        Case &H42D: LangIdent = "Basque"
        Case &H423: LangIdent = "Belarussian"
        Case &H445: LangIdent = "Bengali"
        Case &H402: LangIdent = "Bulgarian"
        Case &H455: LangIdent = "Burmese"
        Case &H403: LangIdent = "Catalan"
        Case &H404: LangIdent = "Chinese (Taiwan)"
        Case &H804: LangIdent = "Chinese (PRC)"
        Case &HC04: LangIdent = "Chinese (Hong Kong SAR, PRC)"
        Case &H1004: LangIdent = "Chinese (Singapore)"
        Case &H1404: LangIdent = "Chinese (Macau SAR)"
        Case &H41A: LangIdent = "Croatian"
        Case &H405: LangIdent = "Czech"
        Case &H406: LangIdent = "Danish"
        Case &H465: LangIdent = "Divehi"
        Case &H413: LangIdent = "Dutch (Netherlands)"
        Case &H813: LangIdent = "Dutch (Belgium)"
        Case &H409: LangIdent = "English (United States)"
        Case &H809: LangIdent = "English (United Kingdom)"
        Case &HC09: LangIdent = "English (Australian)"
        Case &H1009: LangIdent = "English (Canadian)"
        Case &H1409: LangIdent = "English (New Zealand)"
        Case &H1809: LangIdent = "English (Ireland)"
        Case &H1C09: LangIdent = "English (South Africa)"
        Case &H2009: LangIdent = "English (Jamaica)"
        Case &H2409: LangIdent = "English (Caribbean)"
        Case &H2809: LangIdent = "English (Belize)"
        Case &H2C09: LangIdent = "English (Trinidad)"
        Case &H3009: LangIdent = "English (Zimbabwe)"
        Case &H3409: LangIdent = "English (Philippines)"
        Case &H425: LangIdent = "Estonian"
        Case &H438: LangIdent = "Faeroese"
        Case &H429: LangIdent = "Farsi"
        Case &H40B: LangIdent = "Finnish"
        Case &H40C: LangIdent = "French (Standard)"
        Case &H80C: LangIdent = "French (Belgian)"
        Case &HC0C: LangIdent = "French (Canadian)"
        Case &H100C: LangIdent = "French (Switzerland)"
        Case &H140C: LangIdent = "French (Luxembourg)"
        Case &H180C: LangIdent = "French (Monaco)"
        Case &H456: LangIdent = "Galician"
        Case &H43C: LangIdent = "Gaelic - Scotland"
        Case &H437: LangIdent = "Georgian"
        Case &H407: LangIdent = "German (Standard)"
        Case &H807: LangIdent = "German (Switzerland)"
        Case &HC07: LangIdent = "German (Austria)"
        Case &H1007: LangIdent = "German (Luxembourg)"
        Case &H1407: LangIdent = "German (Liechtenstein)"
        Case &H408: LangIdent = "Greek"
        Case &H447: LangIdent = "Gujarati"
        Case &H40D: LangIdent = "Hebrew"
        Case &H439: LangIdent = "Hindi"
        Case &H40E: LangIdent = "Hungarian"
        Case &H40F: LangIdent = "Icelandic"
        Case &H421: LangIdent = "Indonesian"
        Case &H410: LangIdent = "Italian (Standard)"
        Case &H810: LangIdent = "Italian (Switzerland)"
        Case &H411: LangIdent = "Japanese"
        Case &H44B: LangIdent = "Kannada"
        Case &H860: LangIdent = "Kashmiri (India)"
        Case &H43F: LangIdent = "Kazakh"
        Case &H457: LangIdent = "Konkani"
        Case &H412: LangIdent = "Korean"
        Case &H812: LangIdent = "Korean (Johab)"
        Case &H440: LangIdent = "Kyrgyz"
        Case &H426: LangIdent = "Latvian"
        Case &H427: LangIdent = "Lithuanian"
        Case &H827: LangIdent = "Lithuanian (Classic)"
        Case &H42F: LangIdent = "Macedonian"
        Case &H43E: LangIdent = "Malay (Malaysian)"
        Case &H83E: LangIdent = "Malay (Brunei Darussalam)"
        Case &H44C: LangIdent = "Malayalam"
        Case &H43A: LangIdent = "Maltese"
        Case &H458: LangIdent = "Manipuri"
        Case &H44E: LangIdent = "Marathi"
        Case &H450: LangIdent = "Mongolian"
        Case &H861: LangIdent = "Nepali (India)"
        Case &H414: LangIdent = "Norwegian (Bokmal)"
        Case &H814: LangIdent = "Norwegian (Nynorsk)"
        Case &H448: LangIdent = "Oriya"
        Case &H415: LangIdent = "Polish"
        Case &H416: LangIdent = "Portuguese (Brazil)"
        Case &H816: LangIdent = "Portuguese (Standard)"
        Case &H446: LangIdent = "Punjabi"
        Case &H417: LangIdent = "Raeto-Romance"
        Case &H418: LangIdent = "Romanian"
        Case &H818: LangIdent = "Romanian - Moldova"
        Case &H419: LangIdent = "Russian"
        Case &H819: LangIdent = "Russian - Moldova"
        Case &H44F: LangIdent = "Sanskrit"
        Case &HC1A: LangIdent = "Serbian (Cyrillic)"
        Case &H81A: LangIdent = "Serbian (Latin)"
        Case &H459: LangIdent = "Sindhi"
        Case &H41B: LangIdent = "Slovak"
        Case &H424: LangIdent = "Slovenian"
        Case &H42E: LangIdent = "Sorbian"
        Case &H40A: LangIdent = "Spanish (Traditional Sort)"
        Case &H80A: LangIdent = "Spanish (Mexican)"
        Case &HC0A: LangIdent = "Spanish (Modern Sort)"
        Case &H100A: LangIdent = "Spanish (Guatemala)"
        Case &H140A: LangIdent = "Spanish (Costa Rica)"
        Case &H180A: LangIdent = "Spanish (Panama)"
        Case &H1C0A: LangIdent = "Spanish (Dominican Republic)"
        Case &H200A: LangIdent = "Spanish (Venezuela)"
        Case &H240A: LangIdent = "Spanish (Colombia)"
        Case &H280A: LangIdent = "Spanish (Peru)"
        Case &H2C0A: LangIdent = "Spanish (Argentina)"
        Case &H300A: LangIdent = "Spanish (Ecuador)"
        Case &H340A: LangIdent = "Spanish (Chile)"
        Case &H380A: LangIdent = "Spanish (Uruguay)"
        Case &H3C0A: LangIdent = "Spanish (Paraguay)"
        Case &H400A: LangIdent = "Spanish (Bolivia)"
        Case &H440A: LangIdent = "Spanish (El Salvador)"
        Case &H480A: LangIdent = "Spanish (Honduras)"
        Case &H4C0A: LangIdent = "Spanish (Nicaragua)"
        Case &H500A: LangIdent = "Spanish (Puerto Rico)"
        Case &H430: LangIdent = "Sutu"
        Case &H441: LangIdent = "Swahili (Kenya)"
        Case &H41D: LangIdent = "Swedish"
        Case &H81D: LangIdent = "Swedish (Finland)"
        Case &H45A: LangIdent = "Syriac"
        Case &H449: LangIdent = "Tamil"
        Case &H444: LangIdent = "Tatar (Tatarstan)"
        Case &H44A: LangIdent = "Telugu"
        Case &H41E: LangIdent = "Thai"
        Case &H431: LangIdent = "Tsonga"
        Case &H41F: LangIdent = "Turkish"
        Case &H422: LangIdent = "Ukrainian"
        Case &H420: LangIdent = "Urdu (Pakistan)"
        Case &H820: LangIdent = "Urdu (India)"
        Case &H443: LangIdent = "Uzbek (Latin)"
        Case &H843: LangIdent = "Uzbek (Cyrillic)"
        Case &H42A: LangIdent = "Vietnamese"
        Case &H434: LangIdent = "Xhosa"
        Case &H43D: LangIdent = "Yiddish"
        Case &H435: LangIdent = "Zulu"
        Case Else: LangIdent = "Unknown " & lCode
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\LangIdent")
Resume Next
End Function

Public Sub ListView_Clear(ByRef lvwListView As ListView)
On Error GoTo VB_Error

    Dim lIncrement As Long
    For lIncrement = 1 To lvwListView.ListItems.Count
        Call lvwListView.ListItems.Remove(1)
    Next lIncrement
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\ListView_Clear")
Resume Next
End Sub

Public Function LocaleInfo_Get(ByVal lLocale As Long, ByVal LCType As Long) As String
On Error GoTo VB_Error
    
    Dim sBuffer As String
    Dim lBufferLen As Long
    
    If GetLocaleInfo(lLocale, LCType, 0&, lBufferLen) = 0 Then If Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then Call Error_API(Err.LastDllError, sLocation & "\LocaleInfo_Get", "GetLocaleInfo")
    
    sBuffer = String$(256, lBufferLen)
    
    If GetLocaleInfo(lLocale, LCType, sBuffer, Len(sBuffer)) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\LocaleInfo_Get", "GetLocaleInfo")
    
    LocaleInfo_Get = Str_NullTerm_Fix(sBuffer)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\LocaleInfo_Get")
Resume Next
End Function

Public Function PerformanceCounter() As Double
On Error GoTo VB_Error

    Dim LARGE_INTEGER As LARGE_INTEGER
    
    If QueryPerformanceCounter(LARGE_INTEGER) = False Then Call Error_API(Err.LastDllError, sLocation & "\PerformanceCounter", "QueryPerformanceCounter")
    PerformanceCounter = int32x32_int64(LARGE_INTEGER.LowPart, LARGE_INTEGER.HighPart)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\PerformanceCounter")
Resume Next
End Function

Public Function SystemPowerState(ByVal lValue As Long) As String
On Error GoTo VB_Error
    
    Select Case lValue
        Case 0: SystemPowerState = "Unspecified"
        Case 1: SystemPowerState = "Working"
        Case 2: SystemPowerState = "Sleeping1"
        Case 3: SystemPowerState = "Sleeping2"
        Case 4: SystemPowerState = "Sleeping3"
        Case 5: SystemPowerState = "Hibernate"
        Case 6: SystemPowerState = "Shutdown"
        Case Else: SystemPowerState = "Unknown " & lValue
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\SystemPowerState")
Resume Next
End Function

Public Function WindowText_Get(ByVal hwnd As Long) As String
On Error GoTo VB_Error

    Dim sWindowTitle As String
    Dim lRetValue As Long
    
    sWindowTitle = String$(GetWindowTextLength(hwnd) + 1, 0)
    lRetValue = GetWindowText(hwnd, sWindowTitle, Len(sWindowTitle))
    
    If lRetValue = 0 Then
        If Err.LastDllError <> 0 Then Call Error_API(Err.LastDllError, sLocation & "\WindowText_Get", "GetWindowText")
    Else
        WindowText_Get = Str_NullTerm_Fix(Left$(sWindowTitle, lRetValue))
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\WindowText_Get")
Resume Next
End Function

Public Function WinVersion(ByVal lWindows As Long, ByVal lNT As Long, ByVal bRequired As Boolean) As Boolean
On Error GoTo VB_Error

    Select Case lWinID
        Case VER_PLATFORM_WIN32_WINDOWS
            If lWindows = -1 Then
                WinVersion = False
            Else
                If bRequired = True Then
                    If lWindows <= lWinVer Then WinVersion = True
                Else
                    If lWindows > lWinVer Then WinVersion = True
                End If
            End If
        Case VER_PLATFORM_WIN32_NT
            If lNT = -1 Then
                WinVersion = False
            Else
                If bRequired = True Then
                    If lNT <= lWinVer Then WinVersion = True
                Else
                    If lNT > lWinVer Then WinVersion = True
                End If
            End If
    End Select
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\WinVersion")
Resume Next
End Function
