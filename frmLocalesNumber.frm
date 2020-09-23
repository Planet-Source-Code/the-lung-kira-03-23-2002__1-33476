VERSION 5.00
Begin VB.Form frmLocalesNumber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locales - Number"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmLocalesNumber.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDigitGroupSeperator 
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtPositiveSign 
      Height          =   285
      Left            =   4560
      TabIndex        =   24
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtNegativeSign 
      Height          =   285
      Left            =   4560
      TabIndex        =   22
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtNativeDigits 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtDigitGroupingSize 
      Height          =   285
      Left            =   4560
      TabIndex        =   12
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtDecimalSeperator 
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox cboNegativeNumber 
      Height          =   315
      Left            =   4560
      TabIndex        =   20
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CheckBox chkLeadingZeros 
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtDigitSubstitution 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtDigits 
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Top             =   480
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
   Begin VB.ComboBox cboDisplay 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ListBox lstLocales 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5520
      TabIndex        =   25
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblDigitGroupSeperator 
      Caption         =   "Digit Group Seperator"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblPositiveSign 
      Caption         =   "Positive Sign"
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblNegativeSign 
      Caption         =   "Negative Sign"
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblNativeDigits 
      Caption         =   "Native Digits"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblDigitGroupingSize 
      Caption         =   "Digit Grouping Size"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblDecimalSeperator 
      Caption         =   "Decimal Seperator"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblNegativeNumber 
      Caption         =   "Negative Number"
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lblLeadingZeros 
      Caption         =   "Leading Zeros"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblDigitSubstitution 
      Caption         =   "Digit Substitution"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblDigits 
      Caption         =   "Digits"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblDisplay 
      Caption         =   "Display"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
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
Attribute VB_Name = "frmLocalesNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmLocalesNumber"


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

    If Len(txtDecimalSeperator.Text) > 4 Then txtDecimalSeperator.Text = Left$(txtDecimalSeperator.Text, 4)
    If Len(txtDigitGroupSeperator.Text) > 4 Then txtDigitGroupSeperator.Text = Left$(txtDigitGroupSeperator.Text, 4)
    txtDigits.Text = MinMax(Val(txtDigits.Text), 0, 99)
    If Len(txtNegativeSign.Text) > 5 Then txtNegativeSign.Text = Left$(txtNegativeSign.Text, 5)
    If Len(txtPositiveSign.Text) > 5 Then txtPositiveSign.Text = Left$(txtPositiveSign.Text, 5)
    
    
    If lstLocales.ListIndex = -1 Then Exit Sub
    
    
    Dim lLocale As Long
    lLocale = strtoul_(lstLocales.List(lstLocales.ListIndex), 16)
    
    
    If SetLocaleInfo(lLocale, LOCALE_IDIGITS, txtDigits.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_ILZERO, CStr(chkLeadingZeros.value)) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    
    If cboNegativeNumber.ListIndex > -1 Then
        If SetLocaleInfo(lLocale, LOCALE_INEGNUMBER, CStr(cboNegativeNumber.ListIndex)) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    End If
    
    If SetLocaleInfo(lLocale, LOCALE_SDECIMAL, txtDecimalSeperator.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_SMONGROUPING, txtDigitGroupingSize.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_SNEGATIVESIGN, txtNegativeSign.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_SPOSITIVESIGN, txtPositiveSign.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    If SetLocaleInfo(lLocale, LOCALE_STHOUSAND, txtDigitGroupSeperator.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetLocaleInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    lstLocales.Clear
    
    txtDigits.Text = vbNullChar
    txtDigitSubstitution.Text = vbNullChar
    chkLeadingZeros.value = 0
    cboNegativeNumber.ListIndex = -1
    txtDecimalSeperator.Text = vbNullChar
    txtDigitGroupingSize.Text = vbNullChar
    txtNativeDigits.Text = vbNullChar
    txtNegativeSign.Text = vbNullChar
    txtPositiveSign.Text = vbNullChar
    txtDigitGroupSeperator.Text = vbNullChar
    
    
    Dim lFlags As Long
    Select Case cboDisplay.ListIndex
        Case 0: lFlags = LCID_INSTALLED
        Case 1: lFlags = LCID_SUPPORTED
        Case 2: lFlags = LCID_ALTERNATE_SORTS
        Case 3: lFlags = LCID_ALTERNATE_SORTS Or LCID_INSTALLED
        Case 4: lFlags = LCID_ALTERNATE_SORTS Or LCID_SUPPORTED
    End Select
    
    If EnumSystemLocales(AddressOf frmLocalesNumber_EnumLocalesProc, lFlags) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "EnumSystemLocales")
    
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
    With cboNegativeNumber
        .AddItem "Left parenthesis, number, right parenthesis"
        .AddItem "Negative sign, number"
        .AddItem "Negative sign, space, number"
        .AddItem "Number, negative sign"
        .AddItem "Number, space, negative sign"
    End With
    
    
    cboDisplay.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\LocalesNumber", "Display"), 0, 4)
    
    
    If WinVersion(-1, 5000000, True) = False Then
        lblDigitSubstitution.Enabled = False
    End If
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\LocalesNumber", "Display", cboDisplay.ListIndex, REG_DWORD)
    
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
    
    
    txtDigits.Text = LocaleInfo_Get(lLocale, LOCALE_ICURRDIGITS)
    
    If WinVersion(-1, 5000000, True) = True Then
        lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_IDIGITSUBSTITUTION))
        Select Case lReturn
            Case 0: txtDigitSubstitution.Text = "Context"
            Case 1: txtDigitSubstitution.Text = "None/Arabic"
            Case 2: txtDigitSubstitution.Text = "Native"
            Case Else: txtDigitSubstitution.Text = "Unknown " & lReturn
        End Select
    End If
    
    chkLeadingZeros.value = IIf(LocaleInfo_Get(lLocale, LOCALE_ILZERO), 1, 0)
    
    lReturn = Val(LocaleInfo_Get(lLocale, LOCALE_INEGNUMBER))
    Select Case lReturn
        Case 0 To 4: cboNegativeNumber.ListIndex = lReturn
        Case Else: cboNegativeNumber.ListIndex = -1
    End Select
    
    txtDecimalSeperator.Text = LocaleInfo_Get(lLocale, LOCALE_SDECIMAL)
    txtDigitGroupingSize.Text = LocaleInfo_Get(lLocale, LOCALE_SGROUPING)
    txtNativeDigits.Text = LocaleInfo_Get(lLocale, LOCALE_SNATIVEDIGITS)
    txtNegativeSign.Text = LocaleInfo_Get(lLocale, LOCALE_SNEGATIVESIGN)
    txtPositiveSign.Text = LocaleInfo_Get(lLocale, LOCALE_SPOSITIVESIGN)
    txtDigitGroupSeperator.Text = LocaleInfo_Get(lLocale, LOCALE_STHOUSAND)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstLocales_Click")
Resume Next
End Sub
