Attribute VB_Name = "commdlg"
Option Explicit


Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (ByRef lpcc As CHOOSECOLOR_) As Boolean
Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (ByRef lpcf As CHOOSEFONT_) As Boolean
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Boolean
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Boolean


Public Const CC_RGBINIT As Long = &H1
Public Const CC_FULLOPEN As Long = &H2
Public Const CC_PREVENTFULLOPEN As Long = &H4
Public Const CC_SHOWHELP As Long = &H8
Public Const CC_ENABLEHOOK As Long = &H10
Public Const CC_ENABLETEMPLATE As Long = &H20
Public Const CC_ENABLETEMPLATEHANDLE As Long = &H40
Public Const CC_SOLIDCOLOR As Long = &H80
Public Const CC_ANYCOLOR As Long = &H100

Public Const CF_SCREENFONTS As Long = &H1
Public Const CF_PRINTERFONTS As Long = &H2
Public Const CF_BOTH As Long = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP As Long = &H4
Public Const CF_ENABLEHOOK As Long = &H8
Public Const CF_ENABLETEMPLATE As Long = &H10
Public Const CF_ENABLETEMPLATEHANDLE As Long = &H20
Public Const CF_INITTOLOGFONTSTRUCT As Long = &H40
Public Const CF_USESTYLE As Long = &H80
Public Const CF_EFFECTS As Long = &H100
Public Const CF_APPLY As Long = &H200
Public Const CF_ANSIONLY As Long = &H400
Public Const CF_SCRIPTSONLY As Long = CF_ANSIONLY
Public Const CF_NOVECTORFONTS As Long = &H800
Public Const CF_NOOEMFONTS As Long = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS As Long = &H1000
Public Const CF_LIMITSIZE As Long = &H2000
Public Const CF_FIXEDPITCHONLY As Long = &H4000
Public Const CF_WYSIWYG As Long = &H8000
Public Const CF_FORCEFONTEXIST As Long = &H10000
Public Const CF_SCALABLEONLY As Long = &H20000
Public Const CF_TTONLY As Long = &H40000
Public Const CF_NOFACESEL As Long = &H80000
Public Const CF_NOSTYLESEL As Long = &H100000
Public Const CF_NOSIZESEL As Long = &H200000
Public Const CF_SELECTSCRIPT As Long = &H400000
Public Const CF_NOSCRIPTSEL As Long = &H800000
Public Const CF_NOVERTFONTS As Long = &H1000000

Public Const OFN_READONLY As Long = &H1
Public Const OFN_OVERWRITEPROMPT As Long = &H2
Public Const OFN_HIDEREADONLY As Long = &H4
Public Const OFN_NOCHANGEDIR As Long = &H8
Public Const OFN_SHOWHELP As Long = &H10
Public Const OFN_ENABLEHOOK As Long = &H20
Public Const OFN_ENABLETEMPLATE As Long = &H40
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Public Const OFN_NOVALIDATE As Long = &H100
Public Const OFN_ALLOWMULTISELECT As Long = &H200
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_PATHMUSTEXIST As Long = &H800
Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Const OFN_CREATEPROMPT As Long = &H2000
Public Const OFN_SHAREAWARE As Long = &H4000
Public Const OFN_NOREADONLYRETURN As Long = &H8000
Public Const OFN_NOTESTFILECREATE As Long = &H10000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_NOLONGNAMES As Long = &H40000
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_LONGNAMES As Long = &H200000
Public Const OFN_ENABLEINCLUDENOTIFY As Long = &H400000
Public Const OFN_ENABLESIZING As Long = &H800000
Public Const OFN_DONTADDTORECENT As Long = &H2000000
Public Const OFN_FORCESHOWHIDDEN As Long = &H10000000
Public Const OFN_EX_NOPLACESBAR As Long = &H1


Public Type CHOOSECOLOR_
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type CHOOSEFONT_
    lStructSize As Long
    hwndOwner As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
