VERSION 5.00
Begin VB.Form frmLocalesDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locales - Date"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmLocalesDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCentury 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtYearMonthFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtShortDateFormatting 
      Height          =   285
      Left            =   6840
      TabIndex        =   32
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtLongDateFormatting 
      Height          =   285
      Left            =   2400
      TabIndex        =   22
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ListBox lstLongMonthName 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4560
      TabIndex        =   26
      Top             =   1320
      Width           =   4215
   End
   Begin VB.ListBox lstLongDayName 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4560
      TabIndex        =   24
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox txtDateSeperator 
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ListBox lstShortDayName 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4560
      TabIndex        =   34
      Top             =   3240
      Width           =   4215
   End
   Begin VB.ListBox lstShortMonthName 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4560
      TabIndex        =   36
      Top             =   4200
      Width           =   4215
   End
   Begin VB.CheckBox chkDayLeadingZeros 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txtCalendarTypesAvailable 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CheckBox chkMonthLeadingZeros 
      Enabled         =   0   'False
      Height          =   255
      Left            =   8520
      TabIndex        =   28
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txtShortDateFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox cboFirstDayOfWeek 
      Height          =   315
      Left            =   2400
      TabIndex        =   16
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtLongDateFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3960
      Width           =   1935
   End
   Begin VB.ComboBox cboFirstWeekOfYear 
      Height          =   315
      Left            =   2400
      TabIndex        =   18
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ComboBox cboCalendarType 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   7800
      TabIndex        =   39
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox cboDisplay 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox lstLocales 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblYearMonthFormat 
      Caption         =   "Year Month Format"
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblShortDateFormatting 
      Caption         =   "Short Date Formatting"
      Height          =   255
      Left            =   4560
      TabIndex        =   31
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblLongDateFormatting 
      Caption         =   "Long Date Formatting"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblLongMonthName 
      Caption         =   "Long Month Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLongDayName 
      Caption         =   "Long Day Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblDateSeperator 
      Caption         =   "Date Seperator"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblShortDayName 
      Caption         =   "Short Day Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   33
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblShortMonthName 
      Caption         =   "Short Month Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   35
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblCalendarTypesAvailable 
      Caption         =   "Calendar Types Available"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblMonthLeadingZeros 
      Caption         =   "Month Leading Zeros"
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblFirstDayOfWeek 
      Caption         =   "First Day Of Week"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblLongDateFormat 
      Caption         =   "Long Date Format"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblFirstWeekOfYear 
      Caption         =   "First Week Of Year"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lblDayLeadingZeros 
      Caption         =   "Day Leading Zeros"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lblCentury 
      Caption         =   "Century"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblShortDateFormat 
      Caption         =   "Short Date Format"
      Height          =   255
      Left            =   4560
      TabIndex        =   29
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblCalendarType 
      Caption         =   "Calendar Type"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblDisplay 
      Caption         =   "Display"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblLocales 
      Caption         =   "Locales"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmLocalesDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmLocalesDate"


Private Sub cboDisplay_Click()
On Error GoTo VB_Error

    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboDisplay_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    If Len(txtDateSeperator.Text) > 4 Then txtDateSeperator.Text = Left$(txtDateSeperator.Text, 4)
    If Len(txtLongDateFormatting.Text) > 80 Then txtDateSeperator.Text = Left$(txtLongDateFormatting.Text, 80)
    

    If lstLocales.ListIndex = -1 Then Exit Sub
    
    
    Dim lLocale As Long
    lLocale = strtoul_(lstLocales.List(lstLocales.ListIndex), 16)
    
    
    If cboCalendarType.ListIndex > -1 Then
        If SetLocaleInfo(lLocale, LOCALE_ICALENDARTYPE, CStr(cboCalendarType.ListIndex + 1)) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    End If
    If cboFirstWeekOfYear.ListIndex > -1 Then
        If SetLocaleInfo(lLocale, LOCALE_IFIRSTDAYOFWEEK, CStr(cboFirstDayOfWeek.ListIndex)) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    End If
    If cboFirstWeekOfYear.ListIndex > -1 Then
        If SetLocaleInfo(lLocale, LOCALE_IFIRSTWEEKOFYEAR, CStr(cboFirstWeekOfYear.ListIndex)) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    End If
    If SetLocaleInfo(lLocale, LOCALE_SDATE, txtDateSeperator.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_SLONGDATE, txtLongDateFormatting.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_SSHORTDATE, txtShortDateFormatting.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    lstLocales.Clear
    
    cboCalendarType.ListIndex = -1
    txtCentury.Text = vbNullString
    chkDayLeadingZeros.value = 0
    txtShortDateFormat.Text = vbNullString
    cboFirstDayOfWeek.ListIndex = -1
    cboFirstWeekOfYear.ListIndex = -1
    txtLongDateFormat.Text = vbNullString
    chkMonthLeadingZeros.value = 0
    txtCalendarTypesAvailable.Text = vbNullString
    lstShortDayName.Clear
    lstShortMonthName.Clear
    txtDateSeperator.Text = vbNullString
    lstLongDayName.Clear
    lstLongMonthName.Clear
    txtLongDateFormatting.Text = vbNullString
    txtShortDateFormatting.Text = vbNullString
    txtYearMonthFormat.Text = vbNullString
    
        
    Dim lFlags As Long
    Select Case cboDisplay.ListIndex
        Case 0: lFlags = LCID_INSTALLED
        Case 1: lFlags = LCID_SUPPORTED
        Case 2: lFlags = LCID_ALTERNATE_SORTS
        Case 3: lFlags = LCID_ALTERNATE_SORTS Or LCID_INSTALLED
        Case 4: lFlags = LCID_ALTERNATE_SORTS Or LCID_SUPPORTED
    End Select
    
    If EnumSystemLocales(AddressOf frmLocalesDate_EnumLocalesProc, lFlags) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "EnumSystemLocales")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdRefresh_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With cboDisplay
        .AddItem "Installed"
        .AddItem "Supported"
        .AddItem "Alternate Sorts"
        .AddItem "Alternate Sorts + Installed"
        .AddItem "Alternate Sorts + Supported"
    End With
    With cboCalendarType
        .AddItem "Gregorian (localized)"
        .AddItem "Gregorian (English strings always)"
        .AddItem "Year of the Emperor (Japan)"
        .AddItem "Year of Taiwan"
        .AddItem "Tangun Era (Korea)"
        .AddItem "Hijri (Arabic lunar)"
        .AddItem "Thai"
        .AddItem "Hebrew (Lunar)"
        .AddItem "Gregorian Middle East French calendar"
        .AddItem "Gregorian Arabic calendar"
        .AddItem "Gregorian Transliterated English calendar"
        .AddItem "Gregorian Transliterated French calendar"
    End With
    With cboFirstDayOfWeek
        .AddItem "DayName 1"
        .AddItem "DayName 2"
        .AddItem "DayName 3"
        .AddItem "DayName 4"
        .AddItem "DayName 5"
        .AddItem "DayName 6"
        .AddItem "DayName 7"
    End With
    With cboFirstWeekOfYear
        .AddItem "Week containing 1/1 is the first week of that year"
        .AddItem "First full week following 1/1 is the first week of that year"
        .AddItem "First week containing at least four days is the first week of that year"
    End With
    
    
    cboDisplay.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\LocalesDate", "Display"), 0, 4)
    
    
    If WinVersion(-1, 5000000, True) = False Then
        lblYearMonthFormat.Enabled = False
    End If
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesDate", "Display", cboDisplay.ListIndex, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub

Private Sub lstLocales_Click()
On Error GoTo VB_Error

    Dim lLocale As Long
    lLocale = strtoul_(lstLocales.List(lstLocales.ListIndex), 16)
    
    Dim lReturn As Long
    
    
    lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_ICALENDARTYPE)) - 1
    Select Case lReturn
        Case 0 To 11: cboCalendarType.ListIndex = lReturn
        Case Else: cboCalendarType.ListIndex = -1
    End Select
    
    lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_ICENTURY))
    Select Case lReturn
        Case 0: txtCentury.Text = "Abbreviated 2-digit century"
        Case 1: txtCentury.Text = "Full 4-digit century"
        Case Else: txtCentury.Text = "Unknown " & lReturn
    End Select
    
    lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_IDATE))
    Select Case lReturn
        Case 0: txtShortDateFormat.Text = "Month-Day-Year"
        Case 1: txtShortDateFormat.Text = "Day-Month-Year"
        Case 2: txtShortDateFormat.Text = "Year-Month-Day"
        Case Else: txtShortDateFormat.Text = "Unknown " & lReturn
    End Select
    
    chkDayLeadingZeros.value = IIf(LocaleInfo_Get(lLocale, LOCALE_IDAYLZERO), 1, 0)
    
    lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_IFIRSTDAYOFWEEK))
    Select Case lReturn
        Case 0 To 6: cboFirstDayOfWeek.ListIndex = lReturn
        Case Else: cboFirstDayOfWeek.ListIndex = -1
    End Select
    
    lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_IFIRSTWEEKOFYEAR))
    Select Case lReturn
        Case 0 To 2: cboFirstWeekOfYear.ListIndex = lReturn
        Case Else: cboFirstWeekOfYear.ListIndex = -1
    End Select
    
    lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_ILDATE))
    Select Case lReturn
        Case 0: txtLongDateFormat.Text = "Month-Day-Year"
        Case 1: txtLongDateFormat.Text = "Day-Month-Year"
        Case 2: txtLongDateFormat.Text = "Year-Month-Day"
        Case Else: txtLongDateFormat.Text = "Unknown " & lReturn
    End Select
    
    chkMonthLeadingZeros.value = IIf(LocaleInfo_Get(lLocale, LOCALE_IMONLZERO), 1, 0)
    txtCalendarTypesAvailable.Text = RTrim$(Replace$(LocaleInfo_Get(lLocale, LOCALE_IOPTIONALCALENDAR), vbNullChar, " "))
    
    With lstShortDayName
        .Clear
        .AddItem "1 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVDAYNAME1)
        .AddItem "2 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVDAYNAME2)
        .AddItem "3 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVDAYNAME3)
        .AddItem "4 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVDAYNAME4)
        .AddItem "5 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVDAYNAME5)
        .AddItem "6 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVDAYNAME6)
        .AddItem "7 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVDAYNAME7)
    End With
    With lstShortMonthName
        .Clear
        .AddItem "1  " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME1)
        .AddItem "2  " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME2)
        .AddItem "3  " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME3)
        .AddItem "4  " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME4)
        .AddItem "5  " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME5)
        .AddItem "6  " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME6)
        .AddItem "7  " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME7)
        .AddItem "8  " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME8)
        .AddItem "9  " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME9)
        .AddItem "10 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME10)
        .AddItem "11 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME11)
        .AddItem "12 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME12)
        .AddItem "13 " & LocaleInfo_Get(lLocale, LOCALE_SABBREVMONTHNAME13)
    End With
    
    txtDateSeperator.Text = LocaleInfo_Get(lLocale, LOCALE_SDATE)
    
    With lstLongDayName
        .Clear
        .AddItem "1 " & LocaleInfo_Get(lLocale, LOCALE_SDAYNAME1)
        .AddItem "2 " & LocaleInfo_Get(lLocale, LOCALE_SDAYNAME2)
        .AddItem "3 " & LocaleInfo_Get(lLocale, LOCALE_SDAYNAME3)
        .AddItem "4 " & LocaleInfo_Get(lLocale, LOCALE_SDAYNAME4)
        .AddItem "5 " & LocaleInfo_Get(lLocale, LOCALE_SDAYNAME5)
        .AddItem "6 " & LocaleInfo_Get(lLocale, LOCALE_SDAYNAME6)
        .AddItem "7 " & LocaleInfo_Get(lLocale, LOCALE_SDAYNAME7)
    End With
    With lstLongMonthName
        .Clear
        .AddItem "1  " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME1)
        .AddItem "2  " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME2)
        .AddItem "3  " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME3)
        .AddItem "4  " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME4)
        .AddItem "5  " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME5)
        .AddItem "6  " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME6)
        .AddItem "7  " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME7)
        .AddItem "8  " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME8)
        .AddItem "9  " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME9)
        .AddItem "10 " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME10)
        .AddItem "11 " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME11)
        .AddItem "12 " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME12)
        .AddItem "13 " & LocaleInfo_Get(lLocale, LOCALE_SMONTHNAME13)
    End With
    
    txtLongDateFormatting.Text = LocaleInfo_Get(lLocale, LOCALE_SLONGDATE)
    txtShortDateFormatting.Text = LocaleInfo_Get(lLocale, LOCALE_SSHORTDATE)
    
    If WinVersion(-1, 5000000, True) = True Then
        txtYearMonthFormat.Text = LocaleInfo_Get(lLocale, LOCALE_SYEARMONTH)
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstLocales_Click")
Resume Next
End Sub
