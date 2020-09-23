VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessThreadTimes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Thread Times"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmProcessThreadTimes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMillisecondsT 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtSecondsT 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtMinutesT 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtHoursT 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtDaysT 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   2880
      Width           =   615
   End
   Begin VB.ListBox lstThread 
      Height          =   1620
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtMillisecondsK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtSecondsK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtMinutesK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtHoursK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtDaysK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtMillisecondsU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtSecondsU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtMinutesU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtHoursU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtDaysU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtMillisecondE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   30
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtSecondE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   27
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtMinuteE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   24
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtHourE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   21
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtDayE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   18
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtDayOfWeekE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtMonthE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   12
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtYearE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtMillisecondC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   29
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtSecondC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   26
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtMinuteC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   23
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtHourC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   20
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtDayC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtDayOfWeekC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtMonthC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtYearC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2040
      TabIndex        =   4
      Top             =   4320
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
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   255
      Left            =   6120
      TabIndex        =   48
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblThread 
      Caption         =   "Thread"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblMillisecond 
      Caption         =   "Millisecond"
      Height          =   255
      Left            =   3240
      TabIndex        =   28
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblSecond 
      Caption         =   "Second"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblMinute 
      Caption         =   "Minute"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblHour 
      Caption         =   "Hour"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblDay 
      Caption         =   "Day"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblDayOfWeek 
      Caption         =   "Day of Week"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblMonth 
      Caption         =   "Month"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblCreation 
      Caption         =   "Creation"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblSeconds 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   3240
      TabIndex        =   42
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Minutes"
      Height          =   255
      Left            =   3240
      TabIndex        =   39
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblHours 
      Caption         =   "Hours"
      Height          =   255
      Left            =   3240
      TabIndex        =   36
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblDays 
      Caption         =   "Days"
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblKernel 
      Caption         =   "Kernel"
      Height          =   255
      Left            =   4440
      TabIndex        =   31
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblMilliseconds 
      Caption         =   "Milliseconds"
      Height          =   255
      Left            =   3240
      TabIndex        =   45
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblUser 
      Caption         =   "User"
      Height          =   255
      Left            =   5280
      TabIndex        =   32
      Top             =   2640
      Width           =   615
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
Attribute VB_Name = "frmProcessThreadTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Process() As PROCESSENTRY32
Dim Thread() As THREADENTRY32
Const sLocation As String = "frmProcessThreadTimes"


Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    Call ListView_Clear(lvwProcess)
    lstThread.Clear
    
    txtYearC.Text = vbNullString
    txtMonthC.Text = vbNullString
    txtDayOfWeekC.Text = vbNullString
    txtDayC.Text = vbNullString
    txtHourC.Text = vbNullString
    txtMinuteC.Text = vbNullString
    txtSecondC.Text = vbNullString
    txtMillisecondC.Text = vbNullString
    txtYearE.Text = vbNullString
    txtMonthE.Text = vbNullString
    txtDayOfWeekE.Text = vbNullString
    txtDayE.Text = vbNullString
    txtHourE.Text = vbNullString
    txtMinuteE.Text = vbNullString
    txtSecondE.Text = vbNullString
    txtMillisecondE.Text = vbNullString
    txtMillisecondsK.Text = vbNullString
    txtSecondsK.Text = vbNullString
    txtMinutesK.Text = vbNullString
    txtHoursK.Text = vbNullString
    txtDaysK.Text = vbNullString
    txtMillisecondsU.Text = vbNullString
    txtSecondsU.Text = vbNullString
    txtMinutesU.Text = vbNullString
    txtHoursU.Text = vbNullString
    txtDaysU.Text = vbNullString
    txtMillisecondsT.Text = vbNullString
    txtSecondsT.Text = vbNullString
    txtMinutesT.Text = vbNullString
    txtHoursT.Text = vbNullString
    txtDaysT.Text = vbNullString
    
    
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
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = False Then Call DisableAll
    If Function_Exist("kernel32.dll", "OpenThread") = False Then Call DisableAll
    If Function_Exist("kernel32.dll", "GetProcessTimes") = False Then Call DisableAll
    
    If cmdRefresh.Enabled = True Then Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub lstThread_Click()
On Error GoTo VB_Error
    
    Dim hThread As Long
    hThread = OpenThread(THREAD_QUERY_INFORMATION, False, lstThread.List(lstThread.ListIndex)): If hThread = 0 Then Call Error_API(Err.LastDllError, sLocation & "\lstThread_Click", "OpenThread")
    
    
    Dim ftCreation As FILETIME
    Dim ftExit As FILETIME
    Dim ftKernel As FILETIME
    Dim ftUser As FILETIME
    If GetThreadTimes(hThread, ftCreation, ftExit, ftKernel, ftUser) = False Then Call Error_API(Err.LastDllError, sLocation & "\lstThread_Click", "GetThreadTimes")
    
    
    Dim stCreation As SYSTEMTIME
    Dim stExit As SYSTEMTIME
    If int32x32_int64(ftCreation.dwLowDateTime, ftCreation.dwHighDateTime) <> 0 Then
        If FileTimeToSystemTime(ftCreation, stCreation) = False Then Call Error_API(Err.LastDllError, sLocation & "\lstThread_Click", "FiletimeToSystemTime")
    End If
    If int32x32_int64(ftExit.dwLowDateTime, ftExit.dwHighDateTime) <> 0 Then
        If FileTimeToSystemTime(ftExit, stExit) = False Then Call Error_API(Err.LastDllError, sLocation & "\lstThread_Click", "FiletimeToSystemTime")
    End If
    
    
    With stCreation
        txtYearC.Text = .wYear
        txtMonthC.Text = .wMonth
        txtDayOfWeekC.Text = .wDayOfWeek
        txtDayC.Text = .wDay
        txtHourC.Text = .wHour
        txtMinuteC.Text = .wMinute
        txtSecondC.Text = .wSecond
        txtMillisecondC.Text = .wMilliseconds
    End With
    With stExit
        txtYearE.Text = .wYear
        txtMonthE.Text = .wMonth
        txtDayOfWeekE.Text = .wDayOfWeek
        txtDayE.Text = .wDay
        txtHourE.Text = .wHour
        txtMinuteE.Text = .wMinute
        txtSecondE.Text = .wSecond
        txtMillisecondE.Text = .wMilliseconds
    End With
    
    Dim dKernel As Double
    Dim dUser As Double
    dKernel = int32x32_int64(ftKernel.dwLowDateTime, ftKernel.dwHighDateTime)
    dUser = int32x32_int64(ftUser.dwLowDateTime, ftUser.dwHighDateTime)
    
    Dim TIME_LENGTH As TIME_LENGTH
    
    Call Number_TimeLength(dKernel, TIME_LENGTH)
    With TIME_LENGTH
        txtMillisecondsK.Text = .lMilliseconds
        txtSecondsK.Text = .lSeconds
        txtMinutesK.Text = .lMinutes
        txtHoursK.Text = .lHours
        txtDaysK.Text = .lDays
    End With
    
    Call Number_TimeLength(dUser, TIME_LENGTH)
    With TIME_LENGTH
        txtMillisecondsU.Text = .lMilliseconds
        txtSecondsU.Text = .lSeconds
        txtMinutesU.Text = .lMinutes
        txtHoursU.Text = .lHours
        txtDaysU.Text = .lDays
    End With
    
    Call Number_TimeLength(dKernel + dUser, TIME_LENGTH)
    With TIME_LENGTH
        txtMillisecondsT.Text = .lMilliseconds
        txtSecondsT.Text = .lSeconds
        txtMinutesT.Text = .lMinutes
        txtHoursT.Text = .lHours
        txtDaysT.Text = .lDays
    End With
    
    
    If CloseHandle(hThread) = False Then Call Error_API(Err.LastDllError, sLocation & "\lstThread_Click", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstThread_Click")
Resume Next
End Sub

Private Sub lvwProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo VB_Error
    
    If lvwProcess.SelectedItem Is Nothing Then Exit Sub
    
    
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, lvwProcess.SelectedItem): If hProcess = 0 Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "OpenProcess")
    
    
    Dim ftCreation As FILETIME
    Dim ftExit As FILETIME
    Dim ftKernel As FILETIME
    Dim ftUser As FILETIME
    If GetProcessTimes(hProcess, ftCreation, ftExit, ftKernel, ftUser) = False Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "GetProcessTimes")
    
    
    Dim stCreation As SYSTEMTIME
    Dim stExit As SYSTEMTIME
    If int32x32_int64(ftCreation.dwLowDateTime, ftCreation.dwHighDateTime) <> 0 Then
        If FileTimeToSystemTime(ftCreation, stCreation) = False Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "FiletimeToSystemTime")
    End If
    If int32x32_int64(ftExit.dwLowDateTime, ftExit.dwHighDateTime) <> 0 Then
        If FileTimeToSystemTime(ftExit, stExit) = False Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "FiletimeToSystemTime")
    End If
    
    
    With stCreation
        txtYearC.Text = .wYear
        txtMonthC.Text = .wMonth
        txtDayOfWeekC.Text = .wDayOfWeek
        txtDayC.Text = .wDay
        txtHourC.Text = .wHour
        txtMinuteC.Text = .wMinute
        txtSecondC.Text = .wSecond
        txtMillisecondC.Text = .wMilliseconds
    End With
    With stExit
        txtYearE.Text = .wYear
        txtMonthE.Text = .wMonth
        txtDayOfWeekE.Text = .wDayOfWeek
        txtDayE.Text = .wDay
        txtHourE.Text = .wHour
        txtMinuteE.Text = .wMinute
        txtSecondE.Text = .wSecond
        txtMillisecondE.Text = .wMilliseconds
    End With
    
    Dim dKernel As Double
    Dim dUser As Double
    dKernel = int32x32_int64(ftKernel.dwLowDateTime, ftKernel.dwHighDateTime)
    dUser = int32x32_int64(ftUser.dwLowDateTime, ftUser.dwHighDateTime)
    
    Dim TIME_LENGTH As TIME_LENGTH
    
    Call Number_TimeLength(dKernel, TIME_LENGTH)
    With TIME_LENGTH
        txtMillisecondsK.Text = .lMilliseconds
        txtSecondsK.Text = .lSeconds
        txtMinutesK.Text = .lMinutes
        txtHoursK.Text = .lHours
        txtDaysK.Text = .lDays
    End With
    
    Call Number_TimeLength(dUser, TIME_LENGTH)
    With TIME_LENGTH
        txtMillisecondsU.Text = .lMilliseconds
        txtSecondsU.Text = .lSeconds
        txtMinutesU.Text = .lMinutes
        txtHoursU.Text = .lHours
        txtDaysU.Text = .lDays
    End With
    
    Call Number_TimeLength(dKernel + dUser, TIME_LENGTH)
    With TIME_LENGTH
        txtMillisecondsT.Text = .lMilliseconds
        txtSecondsT.Text = .lSeconds
        txtMinutesT.Text = .lMinutes
        txtHoursT.Text = .lHours
        txtDaysT.Text = .lDays
    End With
    
    
    If CloseHandle(hProcess) = False Then Call Error_API(Err.LastDllError, sLocation & "\lvwProcess_ItemClick", "CloseHandle")
    
    
    lstThread.Clear
    
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


Private Sub DisableAll()
On Error GoTo VB_Error
    
    lblProcess.Enabled = False
    lvwProcess.Enabled = False
    lblThread.Enabled = False
    lstThread.Enabled = False
    cmdRefresh.Enabled = False
    
    lblYear.Enabled = False
    lblMonth.Enabled = False
    lblDayOfWeek.Enabled = False
    lblDay.Enabled = False
    lblHour.Enabled = False
    lblMinute.Enabled = False
    lblSecond.Enabled = False
    lblMillisecond.Enabled = False
    lblMilliseconds.Enabled = False
    lblSeconds.Enabled = False
    lblMinutes.Enabled = False
    lblHours.Enabled = False
    lblDays.Enabled = False
        
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\DisableAll")
Resume Next
End Sub
