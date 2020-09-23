Attribute VB_Name = "sfc"
Option Explicit


Public Declare Function SfcGetNextProtectedFile Lib "sfc.dll" (ByVal RpcHandle As Long, ProtFileData As PROTECTED_FILE_DATA) As Long


Public Const SFC_DISABLE_NORMAL As Long = 0
Public Const SFC_DISABLE_ASK As Long = 1
Public Const SFC_DISABLE_ONCE As Long = 2
Public Const SFC_DISABLE_SETUP As Long = 3
Public Const SFC_DISABLE_NOPOPUPS As Long = 4

Public Const SFC_SCAN_NORMAL As Long = 0
Public Const SFC_SCAN_ALWAYS As Long = 1
Public Const SFC_SCAN_ONCE As Long = 2
Public Const SFC_SCAN_IMMEDIATE As Long = 3

Public Const SFC_QUOTA_DEFAULT As Long = 50
Public Const SFC_QUOTA_ALL_FILES As Long = &HFFFFFFFF


Public Type PROTECTED_FILE_DATA
    FileName As String * MAX_PATH
    FileNumber As Long
End Type
