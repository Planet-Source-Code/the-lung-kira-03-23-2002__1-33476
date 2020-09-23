VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmThreads 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Threads"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmThreads.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "Terminate"
      Height          =   350
      Left            =   5160
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdSuspend 
      Caption         =   "Suspend"
      Height          =   350
      Left            =   5160
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Height          =   350
      Left            =   5160
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtSuspendCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtIdealProcessor 
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Text            =   "1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtUsage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtDeltaPriority 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox lstThread 
      Height          =   1620
      Left            =   120
      TabIndex        =   4
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
   Begin VB.Label lblSuspendCount 
      Caption         =   "Suspend Count"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblIdealProcessor 
      Caption         =   "Ideal Processor"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblUsage 
      Caption         =   "Usage"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblDeltaPriority 
      Caption         =   "Delta Priority"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblPriority 
      Caption         =   "Priority"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblThread 
      Caption         =   "Thread ID"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
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
Attribute VB_Name = "frmThreads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Process() As PROCESSENTRY32
Dim Thread() As THREADENTRY32
Const sLocation As String = "frmThreads"


Private Sub cmdApply_Click()
On Error GoTo VB_Error

    txtIdealProcessor.Text = MinMax(Val(txtIdealProcessor.Text), 1, MAXIMUM_PROCESSORS)
    
    
    If lstThread.ListIndex = -1 Then Exit Sub
    
    Dim hThread As Long
    Dim lPriority As Long
    
    hThread = OpenThread(THREAD_SET_INFORMATION, False, lstThread.List(lstThread.ListIndex)): If hThread = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "OpenThread")
    
    
    If cboPriority.ListIndex > -1 Then
        Select Case cboPriority.ListIndex
            Case 0: lPriority = THREAD_PRIORITY_TIME_CRITICAL
            Case 1: lPriority = THREAD_PRIORITY_HIGHEST
            Case 2: lPriority = THREAD_PRIORITY_ABOVE_NORMAL
            Case 3: lPriority = THREAD_PRIORITY_NORMAL
            Case 4: lPriority = THREAD_PRIORITY_BELOW_NORMAL
            Case 5: lPriority = THREAD_PRIORITY_LOWEST
            Case 6: lPriority = THREAD_PRIORITY_IDLE
        End Select
        
        If SetThreadPriority(hThread, lPriority) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetThreadPriority")
    End If
    
    If Function_Exist("kernel32.dll", "SetThreadIdealProcessor") = True Then
        If SetThreadIdealProcessor(hThread, CLng(txtIdealProcessor.Text)) = -1 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetThreadIdealProcessor")
    End If
    
        
    If CloseHandle(hThread) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error

    Call ListView_Clear(lvwProcess)
    lstThread.Clear
    
    cboPriority.ListIndex = -1
    txtDeltaPriority.Text = vbNullString
    txtUsage.Text = vbNullString
    txtSuspendCount.Text = vbNullString
    
    
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

Private Sub cmdResume_Click()
On Error GoTo VB_Error

    If lstThread.ListIndex = -1 Then Exit Sub
    
    Dim hThread As Long
    Dim lSuspendCount As Long
    
    hThread = OpenThread(THREAD_SUSPEND_RESUME, False, lstThread.List(lstThread.ListIndex)): If hThread = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdResume_Click", "OpenThread")
    
    lSuspendCount = ResumeThread(hThread): If lSuspendCount = -1 Then Call Error_API(Err.LastDllError, sLocation & "\cmdResume_Click", "ResumeThread")
    txtSuspendCount.Text = lSuspendCount
    
    If CloseHandle(hThread) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdResume_Click", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdResume_Click")
Resume Next
End Sub

Private Sub cmdSuspend_Click()
On Error GoTo VB_Error

    If lstThread.ListIndex = -1 Then Exit Sub
    
    Dim hThread As Long
    Dim lSuspendCount As Long
    
    hThread = OpenThread(THREAD_SUSPEND_RESUME, False, lstThread.List(lstThread.ListIndex)): If hThread = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdSuspend_Click", "OpenThread")
    
    lSuspendCount = SuspendThread(hThread): If lSuspendCount = -1 Then Call Error_API(Err.LastDllError, sLocation & "\cmdSuspend_Click", "SuspendThread")
    txtSuspendCount.Text = lSuspendCount
    
    If CloseHandle(hThread) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdSuspend_Click", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdSuspend_Click")
Resume Next
End Sub

Private Sub cmdTerminate_Click()
On Error GoTo VB_Error

    If lstThread.ListIndex = -1 Then Exit Sub
    
    Dim hThread As Long
    Dim lExitCode As Long
    
    hThread = OpenThread(THREAD_QUERY_INFORMATION Or THREAD_TERMINATE, False, lstThread.List(lstThread.ListIndex)): If hThread = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdTerminate_Click", "OpenThread")
    If GetExitCodeThread(hThread, lExitCode) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdTerminate_Click", "GetExitCodeThread")
    If TerminateThread(hThread, lExitCode) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdTerminate_Click", "TerminateThread")
    
    If CloseHandle(hThread) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdTerminate_Click", "CloseHandle")
    
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
        .AddItem "Time Critical"
        .AddItem "Highest"
        .AddItem "Above Normal"
        .AddItem "Normal"
        .AddItem "Below Normal"
        .AddItem "Lowest"
        .AddItem "Idle"
    End With
    
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = False Then
        lblProcess.Enabled = False
        lvwProcess.Enabled = False
        cmdRefresh.Enabled = False
        lblThread.Enabled = False
        lstThread.Enabled = False
        lblDeltaPriority.Enabled = False
        txtDeltaPriority.Enabled = False
        lblUsage.Enabled = False
        txtUsage.Enabled = False
    Else
        Call cmdRefresh_Click
    End If
    If Function_Exist("kernel32.dll", "OpenThread") = False Then
        lblPriority.Enabled = False
        cboPriority.Enabled = False
        lblIdealProcessor.Enabled = False
        txtIdealProcessor.Enabled = False
        
        cmdApply.Enabled = False
        cmdResume.Enabled = False
        cmdSuspend.Enabled = False
        cmdTerminate.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub lstThread_Click()
On Error GoTo VB_Error

    If Function_Exist("kernel32.dll", "OpenThread") = True Then
        Dim hThread As Long
        hThread = OpenThread(THREAD_QUERY_INFORMATION, False, lstThread.List(lstThread.ListIndex)): If hThread = 0 Then Call Error_API(Err.LastDllError, sLocation & "\lstThread_Click", "OpenThread")
        
        
        Dim lValue As Long
        lValue = GetThreadPriority(hThread)
        Select Case lValue
            Case THREAD_PRIORITY_LOWEST: cboPriority.ListIndex = 0
            Case THREAD_PRIORITY_BELOW_NORMAL: cboPriority.ListIndex = 1
            Case THREAD_PRIORITY_NORMAL: cboPriority.ListIndex = 2
            Case THREAD_PRIORITY_HIGHEST: cboPriority.ListIndex = 3
            Case THREAD_PRIORITY_ABOVE_NORMAL: cboPriority.ListIndex = 4
            Case THREAD_PRIORITY_ERROR_RETURN: cboPriority.ListIndex = 5
            Case THREAD_PRIORITY_TIME_CRITICAL: cboPriority.ListIndex = 6
            Case THREAD_PRIORITY_IDLE: cboPriority.ListIndex = 7
            Case THREAD_PRIORITY_ERROR_RETURN: Call Error_API(Err.LastDllError, sLocation & "\lstThread_Click", "GetThreadPriority")
            Case Else: cboPriority.Text = lValue
        End Select
        
        If CloseHandle(hThread) = False Then Call Error_API(Err.LastDllError, sLocation & "\lstThread_Click", "CloseHandle")
    Else
        Select Case Thread(lstThread.ListIndex).tpBasePri
            Case THREAD_PRIORITY_LOWEST: cboPriority.ListIndex = 0
            Case THREAD_PRIORITY_BELOW_NORMAL: cboPriority.ListIndex = 1
            Case THREAD_PRIORITY_NORMAL: cboPriority.ListIndex = 2
            Case THREAD_PRIORITY_HIGHEST: cboPriority.ListIndex = 3
            Case THREAD_PRIORITY_ABOVE_NORMAL: cboPriority.ListIndex = 4
            Case THREAD_PRIORITY_ERROR_RETURN: cboPriority.ListIndex = 5
            Case THREAD_PRIORITY_TIME_CRITICAL: cboPriority.ListIndex = 6
            Case THREAD_PRIORITY_IDLE: cboPriority.ListIndex = 7
            Case Else: cboPriority.Text = Thread(lstThread.ListIndex).tpBasePri
        End Select
    End If
    
    
    With Thread(lstThread.ListIndex)
        txtDeltaPriority.Text = .tpDeltaPri
        txtUsage.Text = FormatNumber(int32_uint32(.cntUsage), 0, , , True)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstThread_Click")
Resume Next
End Sub

Private Sub lvwProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo VB_Error
    
    If lvwProcess.SelectedItem Is Nothing Then Exit Sub
    
    
    lstThread.Clear
    
    cboPriority.ListIndex = -1
    txtDeltaPriority.Text = vbNullString
    txtUsage.Text = vbNullString
    txtSuspendCount.Text = vbNullString
    
    
    Dim lCount As Long
    
    lCount = Thread32_Enum(Thread(), lvwProcess.SelectedItem)
    
    Dim lIncrement As Long
    For lIncrement = 0 To lCount
        If Thread(lIncrement).th32OwnerProcessID = lvwProcess.SelectedItem Then
            lstThread.AddItem Thread(lIncrement).th32ThreadID
        End If
    Next lIncrement
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lvwProcess_ItemClick")
Resume Next
End Sub
