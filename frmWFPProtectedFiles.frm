VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWFPProtectedFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows File Protection - Protected Files"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "frmWFPProtectedFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwProtectedFiles 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmWFPProtectedFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmWFPProtectedFiles"


Private Sub Form_Load()
On Error GoTo VB_Error

    lvwProtectedFiles.ColumnHeaders.Add , , "Files"
    
    If Function_Exist("sfc.dll", "SfcGetNextProtectedFile") = True Then
        Dim PROTECTED_FILE_DATA As PROTECTED_FILE_DATA
        PROTECTED_FILE_DATA.FileNumber = 0
        If SfcGetNextProtectedFile(0&, PROTECTED_FILE_DATA) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SfcGetNextProtectedFile")
        
        lvwProtectedFiles.ListItems.Add , , Unicode_Ascii(PROTECTED_FILE_DATA.FileName, 0)
        
        
        With PROTECTED_FILE_DATA
            Do
                PROTECTED_FILE_DATA.FileNumber = .FileNumber + 1
                If SfcGetNextProtectedFile(0&, PROTECTED_FILE_DATA) = 0 Then
                    If Err.LastDllError <> ERROR_NO_MORE_FILES Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SfcGetNextProtectedFile")
                    Exit Do
                End If
                
                lvwProtectedFiles.ListItems.Add , , Unicode_Ascii(.FileName, 0)
                
                If bShutdown = True Then Exit Do
            Loop
        End With
        
        lvwProtectedFiles.Sorted = True
    Else
        lvwProtectedFiles.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
