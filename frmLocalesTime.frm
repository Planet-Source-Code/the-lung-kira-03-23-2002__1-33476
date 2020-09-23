VERSION 5.00
Begin VB.Form frmLocalesTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locales - Time"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmLocalesTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTimeMarkerUse 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtTimeFormatting 
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtTimeSeperator 
      Height          =   285
      Left            =   4560
      TabIndex        =   20
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtPMDesignator 
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtAMDesignator 
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox chkHourLeadingZeros 
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtTimeMarkerPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5520
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.ComboBox cboTimeFormat 
      Height          =   315
      Left            =   4560
      TabIndex        =   12
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ListBox lstLocales 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox cboDisplay 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblTimeFormatting 
      Caption         =   "Time Formatting"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblTimeSeperator 
      Caption         =   "Time Seperator"
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblPMDesignator 
      Caption         =   "PM Designator"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblAMDesignator 
      Caption         =   "AM Designator"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblHourLeadingZeros 
      Caption         =   "Hour Leading Zeros"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblTimeMarkerUse 
      Caption         =   "Time Marker Use"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblTimeMarkerPosition 
      Caption         =   "Time Marker Position"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblTimeFormat 
      Caption         =   "Time Format"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblLocales 
      Caption         =   "Locales"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblDisplay 
      Caption         =   "Display"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmLocalesTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmLocalesTime"


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

    If Len(txtAMDesignator.Text) > 9 Then txtAMDesignator.Text = Left$(txtAMDesignator.Text, 9)
    If Len(txtPMDesignator.Text) > 9 Then txtPMDesignator.Text = Left$(txtPMDesignator.Text, 9)
    If Len(txtTimeFormatting.Text) > 80 Then txtTimeFormatting.Text = Left$(txtTimeFormatting.Text, 80)
    If Len(txtTimeSeperator.Text) > 4 Then txtTimeSeperator.Text = Left$(txtTimeSeperator.Text, 4)
    
    
    If lstLocales.ListIndex = -1 Then Exit Sub
    
    
    Dim lLocale As Long
    lLocale = strtoul_(lstLocales.List(lstLocales.ListIndex), 16)
    
    
    If cboTimeFormat.ListIndex > -1 Then
        If SetLocaleInfo(lLocale, LOCALE_ITIME, CStr(cboTimeFormat.ListIndex)) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    End If
    
    If SetLocaleInfo(lLocale, LOCALE_S1159, txtAMDesignator.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_S2359, txtPMDesignator.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_STIME, txtTimeSeperator.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_STIMEFORMAT, txtTimeFormatting.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    lstLocales.Clear
    
    cboTimeFormat.ListIndex = -1
    txtTimeMarkerPosition.Text = vbNullString
    'txtTimeMarkerUse.Text = vbNullString
    chkHourLeadingZeros.value = 0
    txtAMDesignator.Text = vbNullString
    txtPMDesignator.Text = vbNullString
    txtTimeSeperator.Text = vbNullString
    txtTimeFormatting.Text = vbNullString
    
    
    Dim lFlags As Long
    Select Case cboDisplay.ListIndex
        Case 0: lFlags = LCID_INSTALLED
        Case 1: lFlags = LCID_SUPPORTED
        Case 2: lFlags = LCID_ALTERNATE_SORTS
        Case 3: lFlags = LCID_ALTERNATE_SORTS Or LCID_INSTALLED
        Case 4: lFlags = LCID_ALTERNATE_SORTS Or LCID_SUPPORTED
    End Select
    
    If EnumSystemLocales(AddressOf frmLocalesTime_EnumLocalesProc, lFlags) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "EnumSystemLocales")
    
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
    With cboTimeFormat
        .AddItem "AM / PM 12-hour format"
        .AddItem "24-hour format"
    End With
    
    
    cboDisplay.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\LocalesTime", "Display"), 0, 4)
    
    
    lblTimeMarkerUse.Enabled = False
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesTime", "Display", cboDisplay.ListIndex, REG_DWORD)
    
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
    
    
    lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_ITIME))
    Select Case lReturn
        Case 0 To 1: cboTimeFormat.ListIndex = lReturn
        Case Else: cboTimeFormat.ListIndex = -1
    End Select
    
    lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_ITIMEMARKPOSN))
    Select Case lReturn
        Case 0: txtTimeMarkerPosition.Text = "Use as suffix"
        Case 1: txtTimeMarkerPosition.Text = "Use as prefix"
        Case Else: txtTimeMarkerPosition.Text = "Unknown " & lReturn
    End Select
    'Select Case LocaleInfo_Get(lLocale, LOCALE_ITIMEMARKERUSE)
    '    Case 0: txtTimeMarkerUse.Text = "Use with 12-hour clock"
    '    Case 1: txtTimeMarkerUse.Text = "Use with 24-hour clock"
    '    Case 2: txtTimeMarkerUse.Text = "Use with both 12-hour and 24-hour clocks"
    '    Case 3: txtTimeMarkerUse.Text = "Never use"
    '    Case Else: txtTimeMarkerUse.Text = "Unknown"
    'End Select
    
    chkHourLeadingZeros.value = IIf(LocaleInfo_Get(lLocale, LOCALE_ITLZERO), 1, 0)
    txtAMDesignator.Text = LocaleInfo_Get(lLocale, LOCALE_S1159)
    txtPMDesignator.Text = LocaleInfo_Get(lLocale, LOCALE_S2359)
    txtTimeSeperator.Text = LocaleInfo_Get(lLocale, LOCALE_STIME)
    txtTimeFormatting.Text = LocaleInfo_Get(lLocale, LOCALE_STIMEFORMAT)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstLocales_Click")
Resume Next
End Sub
