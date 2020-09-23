VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSharedFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shared Files"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmSharedFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   6240
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Del Entry"
      Height          =   350
      Left            =   5160
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwLocation 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
Attribute VB_Name = "frmSharedFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmSharedFiles"


Private Sub cmdDelete_Click()
On Error GoTo VB_Error
    
    If lvwLocation.SelectedItem Is Nothing Then Exit Sub
    
    Call Reg_DeleteValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs", lvwLocation.SelectedItem)
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdDelete_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    ListView_Clear lvwLocation
    
    Dim sValueName() As String
    Dim sData() As String
    Dim lDataType() As Long
    Dim lCount As Long
    
    lCount = Reg_EnumValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs", sValueName(), sData(), lDataType())
    
    Dim lIncrement As Long
    For lIncrement = 0 To lCount - 1
        If RTrim$(sValueName(lIncrement)) <> vbNullString Then
            lvwLocation.ListItems.Add(, , sValueName(lIncrement)).SubItems(1) = File_Exist(sValueName(lIncrement))
        End If
    Next lIncrement
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdRefresh_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwLocation.ColumnHeaders
        .Add , , "Location"
        .Add , , "Exist"
    End With
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
