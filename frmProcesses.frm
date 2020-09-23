VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcesses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processes"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmProcesses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "Terminate"
      Height          =   350
      Left            =   2040
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2040
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   975
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox txtUsage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtThreads 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtPrimaryBaseClass 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtParentProcessID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtExpectedVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtAffinityMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtUserObjects 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   22
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtGDIObjects 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   20
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtOtherTransfer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   34
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtWriteTransfer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   32
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtReadTransfer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   30
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtOtherOperation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   28
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtWriteOperation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   26
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtReadOperation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   24
      Top             =   2280
      Width           =   1215
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
   Begin VB.Label lblPriority 
      Caption         =   "Priority"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblUsage 
      Caption         =   "Usage"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblThreads 
      Caption         =   "Threads"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblParentProcessID 
      Caption         =   "Parent Process ID"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblPrimaryBaseClass 
      Caption         =   "Primary Base Class"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblExpectedVersion 
      Caption         =   "Expected Version"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblAffinityMask 
      Caption         =   "Affinity Mask"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblGDIObjects 
      Caption         =   "GDI Objects"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblUserObjects 
      Caption         =   "User Objects"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblOtherTransfer 
      Caption         =   "Other Transfer"
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblWriteTransfer 
      Caption         =   "Write Transfer"
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblReadTransfer 
      Caption         =   "Read Transfer"
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblOtherOperation 
      Caption         =   "Other Operation"
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblWriteOperation 
      Caption         =   "Write Operation"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblReadOperation 
      Caption         =   "Read Operation"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblProcess 
      Caption         =   "Process"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Process() As PROCESSENTRY32
Const sLocation As String = "frmProcesses"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    If lvwProcess.SelectedItem Is Nothing Then Exit Sub
    
    
    If cboPriority.ListIndex > -1 Then
        Dim hProcess As Long
        Dim lPriority As Long
        
        hProcess = OpenProcess(PROCESS_SET_INFORMATION, False, lvwProcess.SelectedItem): If hProcess = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "OpenProcess")
        
        Select Case cboPriority.ListIndex
            Case 0: lPriority = BELOW_NORMAL_PRIORITY_CLASS
            Case 1: lPriority = NORMAL_PRIORITY_CLASS
            Case 2: lPriority = ABOVE_NORMAL_PRIORITY_CLASS
            Case 3: lPriority = REALTIME_PRIORITY_CLASS
            Case 4: lPriority = IDLE_PRIORITY_CLASS
        End Select
        
        If SetPriorityClass(hProcess, lPriority) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetPriorityClass")
        If CloseHandle(hProcess) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "CloseHandle")
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error

    Call ListView_Clear(lvwProcess)
    
    txtAffinityMask.Text = vbNullString
    txtExpectedVersion.Text = vbNullString
    txtParentProcessID.Text = vbNullString
    txtPrimaryBaseClass.Text = vbNullString
    txtThreads.Text = vbNullString
    txtUsage.Text = vbNullString
    txtGDIObjects.Text = vbNullString
    txtUserObjects.Text = vbNullString
    txtReadOperation.Text = vbNullString
    txtWriteOperation.Text = vbNullString
    txtOtherOperation.Text = vbNullString
    txtReadTransfer.Text = vbNullString
    txtWriteTransfer.Text = vbNullString
    txtOtherTransfer.Text = vbNullString
    
    
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

Private Sub cmdTerminate_Click()
On Error GoTo VB_Error
    
    If lvwProcess.SelectedItem Is Nothing Then Exit Sub
    
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, lvwProcess.SelectedItem): If hProcess = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdTerminate_Click", "OpenProcess")
    
    If GetExitCodeProcess(hProcess, lExitCode) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdTerminate_Click", "GetExitCodeProcess")
    If TerminateProcess(hProcess, lExitCode) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdTerminate_Click", "TerminateProcess")
    
    If CloseHandle(hProcess) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdTerminate_Click", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdTerminate_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwProcess.ColumnHeaders
        .Add , , "Process ID"
        .Add , , "Exe Name"
    End With
    With cboPriority
        .AddItem "Below Normal"
        .AddItem "Normal"
        .AddItem "Above Normal"
        .AddItem "Real Time"
        .AddItem "Idle"
    End With
    

    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = False Then
        lblProcess.Enabled = False
        lvwProcess.Enabled = False
        lblPriority.Enabled = False
        cboPriority.Enabled = False
        cmdApply.Enabled = False
        lblAffinityMask.Enabled = False
        lblExpectedVersion.Enabled = False
        lblParentProcessID.Enabled = False
        lblPrimaryBaseClass.Enabled = False
        lblThreads.Enabled = False
        lblUsage.Enabled = False
        cmdRefresh.Enabled = False
        cmdTerminate.Enabled = False
        
        lblGDIObjects.Enabled = False
        lblUserObjects.Enabled = False
        
        lblReadOperation.Enabled = False
        lblWriteOperation.Enabled = False
        lblOtherOperation.Enabled = False
        lblReadTransfer.Enabled = False
        lblWriteTransfer.Enabled = False
        lblOtherTransfer.Enabled = False
    Else
        Call cmdRefresh_Click
    End If
    If Function_Exist("user32.dll", "GetGuiResources") = False Then
        lblGDIObjects.Enabled = False
        lblUserObjects.Enabled = False
    End If
    If Function_Exist("kernel32.dll", "GetProcessIoCounters") = False Then
        lblReadOperation.Enabled = False
        lblWriteOperation.Enabled = False
        lblOtherOperation.Enabled = False
        lblReadTransfer.Enabled = False
        lblWriteTransfer.Enabled = False
        lblOtherTransfer.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub lvwProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo VB_Error
    
    If lvwProcess.SelectedItem Is Nothing Then Exit Sub
    
    
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, lvwProcess.SelectedItem): If hProcess = 0 Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "OpenProcess")
    
    
    Dim lPriorityClass As Long
    lPriorityClass = GetPriorityClass(hProcess)
    Select Case lPriorityClass
        Case BELOW_NORMAL_PRIORITY_CLASS: cboPriority.ListIndex = 0
        Case NORMAL_PRIORITY_CLASS: cboPriority.ListIndex = 1
        Case ABOVE_NORMAL_PRIORITY_CLASS: cboPriority.ListIndex = 2
        Case REALTIME_PRIORITY_CLASS: cboPriority.ListIndex = 3
        Case IDLE_PRIORITY_CLASS: cboPriority.ListIndex = 4
        Case 0: Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "GetPriorityClass")
        Case Else: cboPriority.Text = lPriorityClass
    End Select
    
    
    Dim lProcessAffinityMask As Long
    Dim lSystemAffinityMask As Long
    If GetProcessAffinityMask(hProcess, lProcessAffinityMask, lSystemAffinityMask) = False Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "GetProcessAffinityMask")
    txtAffinityMask.Text = StrReverse(Right$(String$(32, "0") & ltoa_(lProcessAffinityMask, 2), 32))
    
    
    Dim lProcess As Long
    For lProcess = 0 To lvwProcess.ListItems.Count
        If Process(lProcess).th32ProcessID = lvwProcess.SelectedItem Then Exit For
    Next lProcess
    
    With Process(lProcess)
        Dim lExpectedVersion As Long
        lExpectedVersion = GetProcessVersion(.th32ProcessID): If lExpectedVersion = 0 Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "GetProcessVersion")
        txtExpectedVersion.Text = HIWORD(lExpectedVersion) & "." & LOWORD(lExpectedVersion)
        
        txtParentProcessID.Text = .th32ParentProcessID
        txtPrimaryBaseClass.Text = .pcPriClassBase
        txtThreads.Text = FormatNumber(int32_uint32(.cntThreads), 0, , , True)
        txtUsage.Text = FormatNumber(int32_uint32(.cntUsage), 0, , , True)
    End With
    
    
    If Function_Exist("user32.dll", "GetGuiResources") = True Then
        Dim lValue As Long
        
        lValue = GetGuiResources(hProcess, GR_GDIOBJECTS): If lValue = 0 Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "GetGuiResources")
        txtGDIObjects.Text = FormatNumber(int32_uint32(lValue), 0, , , True)
        
        lValue = GetGuiResources(hProcess, GR_USEROBJECTS): If lValue = 0 Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "GetGuiResources")
        txtUserObjects.Text = FormatNumber(int32_uint32(lValue), 0, , , True)
    End If
    
    If Function_Exist("kernel32.dll", "GetProcessIoCounters") = True Then
        Dim IO_COUNTERS As IO_COUNTERS
        If GetProcessIoCounters(hProcess, IO_COUNTERS) = False Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "GetProcessIoCounters")
        
        With IO_COUNTERS
            txtReadOperation.Text = FormatNumber(int32x32_int64(.ReadOperationCount.LowPart, .ReadOperationCount.HighPart), 0, , , True)
            txtWriteOperation.Text = FormatNumber(int32x32_int64(.WriteOperationCount.LowPart, .WriteOperationCount.HighPart), 0, , , True)
            txtOtherOperation.Text = FormatNumber(int32x32_int64(.OtherOperationCount.LowPart, .OtherOperationCount.HighPart), 0, , , True)
            txtReadTransfer.Text = FormatNumber(int32x32_int64(.ReadTransferCount.LowPart, .ReadTransferCount.HighPart), 0, , , True)
            txtWriteTransfer.Text = FormatNumber(int32x32_int64(.WriteTransferCount.LowPart, .WriteTransferCount.HighPart), 0, , , True)
            txtOtherTransfer.Text = FormatNumber(int32x32_int64(.OtherTransferCount.LowPart, .OtherTransferCount.HighPart), 0, , , True)
        End With
    End If
    
    If CloseHandle(hProcess) = False Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lvwProcess_ItemClick")
Resume Next
End Sub
