VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StartUp"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmStartUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   5640
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Del Entry"
      Height          =   350
      Left            =   4560
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwStartUp 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3836
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
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmStartUp"


Private Sub cmdDelete_Click()
On Error GoTo VB_Error
    
    If lvwStartUp.SelectedItem Is Nothing Then Exit Sub
    
    
    Select Case lvwStartUp.SelectedItem.SubItems(2)
        Case "User": Call Reg_DeleteValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\" & lvwStartUp.SelectedItem.SubItems(3), lvwStartUp.SelectedItem)
        Case "Machine": Call Reg_DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\" & lvwStartUp.SelectedItem.SubItems(3), lvwStartUp.SelectedItem)
    End Select
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdDelete_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    Call ListView_Clear(lvwStartUp)
    
    Call ProcessStartup(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "User", "Run")
    Call ProcessStartup(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", "User", "RunOnce")
    Call ProcessStartup(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Machine", "Run")
    Call ProcessStartup(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Runonce", "Machine", "RunOnce")
    Call ProcessStartup(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", "Machine", "RunServices")
    Call ProcessStartup(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce", "Machine", "RunServicesOnce")
    Call ProcessStartup(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce\Setup", "Machine", "RunOnce\Setup")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdRefresh_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwStartUp
        .ColumnHeaders.Add , , "Name"
        .ColumnHeaders.Add , , "Entry"
        .ColumnHeaders.Add , , "Owner"
        .ColumnHeaders.Add , , "Type"
    End With
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub ProcessStartup(ByVal lHkey As Long, ByVal sPath As String, ByVal sSubItem2 As String, ByVal sSubItem3 As String)
On Error GoTo VB_Error

    Dim sValueName() As String
    Dim sData() As String
    Dim lDataType() As Long
    Dim lCount As Long
    Dim lIncrement As Long
    
    lCount = Reg_EnumValue(lHkey, sPath, sValueName(), sData(), lDataType())
    For lIncrement = 0 To lCount - 1
        If RTrim$(sValueName(lIncrement)) <> vbNullString Then
            With lvwStartUp.ListItems.Add(, , sValueName(lIncrement))
                .SubItems(1) = sData(lIncrement)
                .SubItems(2) = sSubItem2
                .SubItems(3) = sSubItem3
            End With
        End If
    Next lIncrement
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\ProcessStartup")
Resume Next
End Sub
