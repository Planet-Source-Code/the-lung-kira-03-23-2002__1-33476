VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModules 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modules"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmModules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtB2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "B"
      Top             =   3000
      Width           =   135
   End
   Begin VB.TextBox txtUsageCount 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtGlobalUsageCount 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtExePath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox txtBaseSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtBaseAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ListBox lstModule 
      Height          =   1620
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwProcess 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2990
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
   Begin VB.Label lblUsageCount 
      Caption         =   "Usage Count"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblGlobalUsageCount 
      Caption         =   "Global Usage Count"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblBaseAddress 
      Caption         =   "Base Address"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblBaseSize 
      Caption         =   "Base Size"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblExePath 
      Caption         =   "Exe Path"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblModule 
      Caption         =   "Module Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblProcess 
      Caption         =   "Process ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Process() As PROCESSENTRY32
Dim Module() As MODULEENTRY32
Const sLocation As String = "frmModules"


Private Sub cmdRefresh_Click()
On Error GoTo VB_Error

    ListView_Clear lvwProcess
    lstModule.Clear
    
    txtBaseAddress.Text = vbNullString
    txtBaseSize.Text = vbNullString
    txtExePath.Text = vbNullString
    txtGlobalUsageCount.Text = vbNullString
    txtUsageCount.Text = vbNullString
    
    
    Dim lCount As Long
    
    lCount = Process32_Enum(Process())
    
    Dim lIncrement As Long
    For lIncrement = 0 To lCount
        lvwProcess.ListItems.Add(, , Process(lIncrement).th32ProcessID).SubItems(1) = Process(lIncrement).szExeFile
    Next lIncrement
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdRefresh_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwProcess.ColumnHeaders
        .Add , , "Process ID"
        .Add , , "Exe Name"
    End With
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = False Then
        lblProcess.Enabled = False
        lvwProcess.Enabled = False
        cmdRefresh.Enabled = False
        lblModule.Enabled = False
        lstModule.Enabled = False
        lblBaseAddress.Enabled = False
        lblBaseSize.Enabled = False
        lblExePath.Enabled = False
        lblGlobalUsageCount.Enabled = False
        lblUsageCount.Enabled = False
    Else
        Call cmdRefresh_Click
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub lstModule_Click()
On Error GoTo VB_Error

    With Module(lstModule.ListIndex)
        txtBaseAddress.Text = .modBaseAddr
        txtBaseSize.Text = FormatNumber(.modBaseSize, 0, , , True)
        txtExePath.Text = .szExePath
        txtGlobalUsageCount.Text = FormatNumber(int32_uint32(.GlblcntUsage), 0, , , True)
        txtUsageCount.Text = FormatNumber(int32_uint32(.ProccntUsage), 0, , , True)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstModule_Click")
Resume Next
End Sub

Private Sub lvwProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo VB_Error
    
    If lvwProcess.SelectedItem Is Nothing Then Exit Sub
    
    
    lstModule.Clear
    
    txtBaseAddress.Text = vbNullString
    txtBaseSize.Text = vbNullString
    txtExePath.Text = vbNullString
    txtGlobalUsageCount.Text = vbNullString
    txtUsageCount.Text = vbNullString
    
    
    Dim lCount As Long
    
    lCount = Module32_Enum(Module(), lvwProcess.SelectedItem)
    
    Dim lIncrement As Long
    Dim lSize As Long
    Dim lUsage As Long
    
    For lIncrement = 0 To lCount
        lstModule.AddItem Module(lIncrement).szModule
    Next lIncrement
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lvwProcess_ItemClick")
Resume Next
End Sub
