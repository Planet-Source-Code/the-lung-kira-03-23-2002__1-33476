VERSION 5.00
Begin VB.Form frmFileTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Time"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmFileTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox drvFileTime 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox txtSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   6975
   End
   Begin VB.DirListBox dirFileTime 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.FileListBox fileFileTime 
      Height          =   2040
      Hidden          =   -1  'True
      Left            =   120
      System          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtMillisecondLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   39
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtMillisecondLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   38
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtMillisecondCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   37
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtSecondLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   35
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtSecondLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   34
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtSecondCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   33
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtMinuteLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   31
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtMinuteLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   30
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtMinuteCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   29
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtHourLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   27
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtHourLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   26
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtHourCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   25
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtDayLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   23
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtDayLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   22
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtDayCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtDayOfWeekLW 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtDayOfWeekLA 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtDayOfWeekCT 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtMonthLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMonthLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMonthCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtYearLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtYearLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtYearCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   6120
      TabIndex        =   40
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblSelected 
      Caption         =   "Selected"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblMillisecond 
      Caption         =   "Millisecond"
      Height          =   255
      Left            =   2520
      TabIndex        =   36
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblSecond 
      Caption         =   "Second"
      Height          =   255
      Left            =   2520
      TabIndex        =   32
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblMinute 
      Caption         =   "Minute"
      Height          =   255
      Left            =   2520
      TabIndex        =   28
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblHour 
      Caption         =   "Hour"
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblDay 
      Caption         =   "Day"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblDayOfWeek 
      Caption         =   "Day of Week"
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblMonth 
      Caption         =   "Month"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblLastWrite 
      Caption         =   "Last Write"
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblLastAccess 
      Caption         =   "Last Access"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblCreation 
      Caption         =   "Creation"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmFileTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmFileTime"


Private Sub cmdApply_Click()
On Error GoTo VB_Error

    If txtSelected.Text = vbNullString Then Exit Sub
    
    
    txtDayCT.Text = MinMax(Val(txtDayCT.Text), 0, 31)
    txtDayLA.Text = MinMax(Val(txtDayLA.Text), 0, 31)
    txtDayLW.Text = MinMax(Val(txtDayLW.Text), 0, 31)
    txtHourCT.Text = MinMax(Val(txtHourCT.Text), 0, 23)
    txtHourLA.Text = MinMax(Val(txtHourLA.Text), 0, 23)
    txtHourLW.Text = MinMax(Val(txtHourLW.Text), 0, 23)
    txtMillisecondCT.Text = MinMax(Val(txtMillisecondCT.Text), 0, 999)
    txtMillisecondLA.Text = MinMax(Val(txtMillisecondLA.Text), 0, 999)
    txtMillisecondLW.Text = MinMax(Val(txtMillisecondLW.Text), 0, 999)
    txtMinuteCT.Text = MinMax(Val(txtMinuteCT.Text), 0, 59)
    txtMinuteLA.Text = MinMax(Val(txtMinuteLA.Text), 0, 59)
    txtMinuteLW.Text = MinMax(Val(txtMinuteLW.Text), 0, 59)
    txtMonthCT.Text = MinMax(Val(txtMonthCT.Text), 0, 12)
    txtMonthLA.Text = MinMax(Val(txtMonthLA.Text), 0, 12)
    txtMonthLW.Text = MinMax(Val(txtMonthLW.Text), 0, 12)
    txtSecondCT.Text = MinMax(Val(txtSecondCT.Text), 0, 59)
    txtSecondLA.Text = MinMax(Val(txtSecondLA.Text), 0, 59)
    txtSecondLW.Text = MinMax(Val(txtSecondLW.Text), 0, 59)
    If WinVersion(-1, 5010000, True) = True Then
        txtYearCT.Text = MinMax(Val(txtYearCT.Text), 1601, 30827)
        txtYearLA.Text = MinMax(Val(txtYearLA.Text), 1601, 30827)
        txtYearLW.Text = MinMax(Val(txtYearLW.Text), 1601, 30827)
    Else
        txtYearCT.Text = MinMax(Val(txtYearCT.Text), 1601, 65535)
        txtYearLA.Text = MinMax(Val(txtYearLA.Text), 1601, 65535)
        txtYearLW.Text = MinMax(Val(txtYearLW.Text), 1601, 65535)
    End If
    
    
    Dim hFile As Long
    
    Dim ftCreationTime As FILETIME
    Dim ftLastAccess As FILETIME
    Dim ftLastWrite As FILETIME
    Dim stCreationTime As SYSTEMTIME
    Dim stLastAccess As SYSTEMTIME
    Dim stLastWrite As SYSTEMTIME
    
    
    hFile = CreateFile(txtSelected.Text, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&): If hFile = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "CreateFile")
    
    
    With stCreationTime
        .wYear = uint16_int16(txtYearCT.Text)
        .wMonth = txtMonthCT.Text
        .wDayOfWeek = txtDayOfWeekCT.Text
        .wDay = txtDayCT.Text
        .wHour = txtHourCT.Text
        .wMinute = txtMinuteCT.Text
        .wSecond = txtSecondCT.Text
        .wMilliseconds = txtMillisecondCT.Text
    End With
    With stLastAccess
        .wYear = uint16_int16(txtYearLA.Text)
        .wMonth = txtMonthLA.Text
        .wDayOfWeek = txtDayOfWeekLA.Text
        .wDay = txtDayLA.Text
        .wHour = txtHourLA.Text
        .wMinute = txtMinuteLA.Text
        .wSecond = txtSecondLA.Text
        .wMilliseconds = txtMillisecondLA.Text
    End With
    With stLastWrite
        .wYear = uint16_int16(txtYearLW.Text)
        .wMonth = txtMonthLW.Text
        .wDayOfWeek = txtDayOfWeekLW.Text
        .wDay = txtDayLW.Text
        .wHour = txtHourLW.Text
        .wMinute = txtMinuteLW.Text
        .wSecond = txtSecondLW.Text
        .wMilliseconds = txtMillisecondLW.Text
    End With
    
    
    If SystemTimeToFileTime(stCreationTime, ftCreationTime) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemTimeToFileTime")
    If SystemTimeToFileTime(stLastAccess, ftLastAccess) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemTimeToFileTime")
    If SystemTimeToFileTime(stLastWrite, ftLastWrite) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemTimeToFileTime")
    
    If SetFileTime(hFile, ftCreationTime, ftLastAccess, ftLastWrite) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetFileTime")
    If CloseHandle(hFile) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub dirFileTime_Change()
On Error GoTo VB_Error

    fileFileTime.Path = dirFileTime.Path
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\dirFileTime_Change")
Resume Next
End Sub

Private Sub drvFileTime_Change()
On Error GoTo VB_Error
    
    dirFileTime.Path = drvFileTime.Drive
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\drvFileTime_Change")
Resume Next
End Sub

Private Sub fileFileTime_Click()
On Error GoTo VB_Error

    txtSelected.Text = Str_BckSlhTerm_Fix(dirFileTime.Path) & "\" & fileFileTime.FileName
    ProcessFileTime txtSelected.Text
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\fileFileTime_Click")
Resume Next
End Sub

Private Sub ProcessFileTime(strFileName As String)
On Error GoTo VB_Error

    Dim hFile As Long
    
    Dim ftCreationTime As FILETIME
    Dim ftLastAccess As FILETIME
    Dim ftLastWrite As FILETIME
    Dim stCreationTime As SYSTEMTIME
    Dim stLastAccess As SYSTEMTIME
    Dim stLastWrite As SYSTEMTIME


    hFile = CreateFile(strFileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&): If hFile = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\ProcessFileTime", "CreateFile")
    
    If GetFileTime(hFile, ftCreationTime, ftLastAccess, ftLastWrite) = False Then Call Error_API(Err.LastDllError, sLocation & "\ProcessFileTime", "GetFileTime")
    
    If FileTimeToSystemTime(ftCreationTime, stCreationTime) = False Then Call Error_API(Err.LastDllError, sLocation & "\ProcessFileTime", "FiletimeToSystemTime")
    If FileTimeToSystemTime(ftLastAccess, stLastAccess) = False Then Call Error_API(Err.LastDllError, sLocation & "\ProcessFileTime", "FiletimeToSystemTime")
    If FileTimeToSystemTime(ftLastWrite, stLastWrite) = False Then Call Error_API(Err.LastDllError, sLocation & "\ProcessFileTime", "FiletimeToSystemTime")
    
    
    With stCreationTime
        txtYearCT.Text = .wYear
        txtMonthCT.Text = .wMonth
        txtDayOfWeekCT.Text = .wDayOfWeek
        txtDayCT.Text = .wDay
        txtHourCT.Text = .wHour
        txtMinuteCT.Text = .wMinute
        txtSecondCT.Text = .wSecond
        txtMillisecondCT.Text = .wMilliseconds
    End With
    With stLastAccess
        txtYearLA.Text = .wYear
        txtMonthLA.Text = .wMonth
        txtDayOfWeekLA.Text = .wDayOfWeek
        txtDayLA.Text = .wDay
        txtHourLA.Text = .wHour
        txtMinuteLA.Text = .wMinute
        txtSecondLA.Text = .wSecond
        txtMillisecondLA.Text = .wMilliseconds
    End With
    With stLastWrite
        txtYearLW.Text = .wYear
        txtMonthLW.Text = .wMonth
        txtDayOfWeekLW.Text = .wDayOfWeek
        txtDayLW.Text = .wDay
        txtHourLW.Text = .wHour
        txtMinuteLW.Text = .wMinute
        txtSecondLW.Text = .wSecond
        txtMillisecondLW.Text = .wMilliseconds
    End With
    
    If CloseHandle(hFile) = False Then Call Error_API(Err.LastDllError, sLocation & "\ProcessFileTime", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\ProcessFileTime")
Resume Next
End Sub
