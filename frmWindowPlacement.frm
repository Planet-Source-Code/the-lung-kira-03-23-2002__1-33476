VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWindowPlacement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Placement"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmWindowPlacement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstThread 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtNormalPositionTop 
      Height          =   285
      Left            =   4680
      TabIndex        =   24
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtNormalPositionBottom 
      Height          =   285
      Left            =   4680
      TabIndex        =   26
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5400
      TabIndex        =   37
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtWindowText 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2040
      TabIndex        =   6
      Top             =   5520
      Width           =   975
   End
   Begin VB.CheckBox chkSetMinimizedPosition 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chkRestoreToMaximized 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox chkAsyncWindowPlacement 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   1200
      Width           =   255
   End
   Begin VB.ComboBox cboShowWindow 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      TabIndex        =   10
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtMinimizedPositionX 
      Height          =   285
      Left            =   4680
      TabIndex        =   29
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtMinimizedPositionY 
      Height          =   285
      Left            =   4680
      TabIndex        =   31
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtMaximizedPositionX 
      Height          =   285
      Left            =   4680
      TabIndex        =   34
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtMaximizedPositionY 
      Height          =   285
      Left            =   4680
      TabIndex        =   36
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtNormalPositionLeft 
      Height          =   285
      Left            =   4680
      TabIndex        =   20
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtNormalPositionRight 
      Height          =   285
      Left            =   4680
      TabIndex        =   22
      Top             =   2760
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwProcess 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2566
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
   Begin MSComctlLib.ListView lvwWindow 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2566
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
   Begin VB.Label lblWindow 
      Caption         =   "Window"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
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
   Begin VB.Label lblThread 
      Caption         =   "Thread"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblMaximizedPositionX 
      Caption         =   "X"
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lblMaximizedPositionY 
      Caption         =   "Y"
      Height          =   255
      Left            =   3240
      TabIndex        =   35
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lblNormalPositionBottom 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblNormalPositionTop 
      Caption         =   "Top"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblNormalPositionRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblNormalPositionLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblWindow_Text 
      Caption         =   "Window Text"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblMinimizedPositionY 
      Caption         =   "Y"
      Height          =   255
      Left            =   3240
      TabIndex        =   30
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblMinimizedPositionX 
      Caption         =   "X"
      Height          =   255
      Left            =   3240
      TabIndex        =   28
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblSetMinimizedPosition 
      Caption         =   "Set Minimized Position"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblRestoreToMaximized 
      Caption         =   "Restore To Maximized"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblAsyncWindowPlacement 
      Caption         =   "Async Window Placement"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblFlags 
      Caption         =   "Flags"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblShowWindow 
      Caption         =   "Show Window"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblMinimizedPosition 
      Caption         =   "Minimized Position"
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblMaximizedPosition 
      Caption         =   "Maximized Position"
      Height          =   255
      Left            =   3240
      TabIndex        =   32
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblNormalPosition 
      Caption         =   "Normal Position"
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frmWindowPlacement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmWindowPlacement"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    If lvwWindow.SelectedItem Is Nothing Then Exit Sub
    
    
    txtMaximizedPositionX.Text = MinMax(Val(txtMaximizedPositionX.Text), -2147483648#, 2147483647)
    txtMaximizedPositionY.Text = MinMax(Val(txtMaximizedPositionY.Text), -2147483648#, 2147483647)
    txtMinimizedPositionX.Text = MinMax(Val(txtMinimizedPositionX.Text), -2147483648#, 2147483647)
    txtMinimizedPositionY.Text = MinMax(Val(txtMinimizedPositionY.Text), -2147483648#, 2147483647)
    txtNormalPositionBottom.Text = MinMax(Val(txtNormalPositionBottom.Text), -2147483648#, 2147483647)
    txtNormalPositionLeft.Text = MinMax(Val(txtNormalPositionLeft.Text), -2147483648#, 2147483647)
    txtNormalPositionRight.Text = MinMax(Val(txtNormalPositionRight.Text), -2147483648#, 2147483647)
    txtNormalPositionTop.Text = MinMax(Val(txtNormalPositionTop.Text), -2147483648#, 2147483647)
    
    
    Dim lHandle As Long
    lHandle = lvwWindow.SelectedItem
    
    If lHandle <> 0 Then
        Dim WINDOWPLACEMENT As WINDOWPLACEMENT
        With WINDOWPLACEMENT
            .Length = Len(WINDOWPLACEMENT)
            
            If WinVersion(-1, 5000000, True) = True Then
                If chkAsyncWindowPlacement.value = 1 Then .flags = .flags Or WPF_ASYNCWINDOWPLACEMENT
            End If
            If chkRestoreToMaximized.value = 1 Then .flags = .flags Or WPF_RESTORETOMAXIMIZED
            If chkSetMinimizedPosition.value = 1 Then .flags = .flags Or WPF_SETMINPOSITION
            
            Select Case cboShowWindow.ListIndex
                Case 0: .showCmd = SW_HIDE
                Case 1: .showCmd = SW_SHOWNORMAL
                Case 2: .showCmd = SW_SHOWMINIMIZED
                Case 3: .showCmd = SW_SHOWMAXIMIZED
                Case 4: .showCmd = SW_SHOWNOACTIVATE
                Case 5: .showCmd = SW_SHOW
                Case 6: .showCmd = SW_MINIMIZE
                Case 7: .showCmd = SW_SHOWMINNOACTIVE
                Case 8: .showCmd = SW_SHOWNA
                Case 9: .showCmd = SW_RESTORE
                Case 10: .showCmd = SW_SHOWDEFAULT
                Case 11: .showCmd = SW_FORCEMINIMIZE
            End Select
            
            .ptMinPosition.X = txtMinimizedPositionX.Text
            .ptMinPosition.Y = txtMinimizedPositionY.Text
            .ptMaxPosition.X = txtMaximizedPositionX.Text
            .ptMaxPosition.Y = txtMaximizedPositionY.Text
            .rcNormalPosition.Left = txtNormalPositionLeft.Text
            .rcNormalPosition.Right = txtNormalPositionRight.Text
            .rcNormalPosition.Top = txtNormalPositionTop.Text
            .rcNormalPosition.Bottom = txtNormalPositionBottom.Text
        End With
        
        If SetWindowPlacement(lHandle, WINDOWPLACEMENT) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetWindowPlacement")
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    Call ListView_Clear(lvwProcess)
    lstThread.Clear
    Call ListView_Clear(lvwWindow)
    
    
    Dim Process() As PROCESSENTRY32
    
    
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
    With lvwWindow.ColumnHeaders
        .Add , , "Handle"
        .Add , , "Window Text"
    End With
    
    With cboShowWindow
        .AddItem "Hide"
        .AddItem "Show Normal"
        .AddItem "Show Minimized"
        .AddItem "Show Maximized"
        .AddItem "Show Not Active"
        .AddItem "Show"
        .AddItem "Minimize"
        .AddItem "Show Minimized Not Activated"
        .AddItem "Show NA"
        .AddItem "Restore"
        .AddItem "Show Default"
        .AddItem "Force Minimize"
    End With
    
    
    If WinVersion(-1, 5000000, True) = False Then
        lblAsyncWindowPlacement.Enabled = False
        chkAsyncWindowPlacement.Enabled = False
    End If
    
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub lstThread_Click()
On Error GoTo VB_Error
    
    Call ListView_Clear(lvwWindow)
    
    If EnumWindows(AddressOf frmWindowPlacement_EnumWindowsProc, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "EnumWindows")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstThread_Click")
Resume Next
End Sub

Private Sub lvwProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo VB_Error
    
    If lvwProcess.SelectedItem Is Nothing Then Exit Sub
    
    
    lstThread.Clear
    Call ListView_Clear(lvwWindow)
    
    
    Dim Thread() As THREADENTRY32
    
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

Private Sub lvwWindow_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo VB_Error
    
    If lvwWindow.SelectedItem Is Nothing Then Exit Sub
    
    
    Dim lHandle As Long
    lHandle = lvwWindow.SelectedItem
    txtWindowText.Text = lvwWindow.SelectedItem.SubItems(1)
    
        
    Dim WINDOWPLACEMENT As WINDOWPLACEMENT
    WINDOWPLACEMENT.Length = Len(WINDOWPLACEMENT)
    If GetWindowPlacement(lHandle, WINDOWPLACEMENT) = False Then Call Error_API(Err.LastDllError, sLocation & "\lvwWindow_ItemClick", "GetWindowPlacement")
    
    With WINDOWPLACEMENT
        If WinVersion(-1, 5000000, True) = True Then
            chkAsyncWindowPlacement.value = IIf(.flags And WPF_ASYNCWINDOWPLACEMENT, 1, 0)
        End If
        
        chkRestoreToMaximized.value = IIf(.flags And WPF_RESTORETOMAXIMIZED, 1, 0)
        chkSetMinimizedPosition.value = IIf(.flags And WPF_SETMINPOSITION, 1, 0)
        
        Select Case .showCmd
            Case SW_HIDE: cboShowWindow.ListIndex = 0
            Case SW_SHOWNORMAL: cboShowWindow.ListIndex = 1
            Case SW_SHOWMINIMIZED: cboShowWindow.ListIndex = 2
            Case SW_SHOWMAXIMIZED: cboShowWindow.ListIndex = 3
            Case SW_SHOWNOACTIVATE: cboShowWindow.ListIndex = 4
            Case SW_SHOW: cboShowWindow.ListIndex = 5
            Case SW_MINIMIZE: cboShowWindow.ListIndex = 6
            Case SW_SHOWMINNOACTIVE: cboShowWindow.ListIndex = 7
            Case SW_SHOWNA: cboShowWindow.ListIndex = 8
            Case SW_RESTORE: cboShowWindow.ListIndex = 9
            Case SW_SHOWDEFAULT: cboShowWindow.ListIndex = 10
            Case SW_FORCEMINIMIZE: cboShowWindow.ListIndex = 11
            Case Else: cboShowWindow.ListIndex = -1
        End Select
        
        txtMinimizedPositionX.Text = .ptMinPosition.X
        txtMinimizedPositionY.Text = .ptMinPosition.Y
        txtMaximizedPositionX.Text = .ptMaxPosition.X
        txtMaximizedPositionY.Text = .ptMaxPosition.Y
        txtNormalPositionLeft.Text = .rcNormalPosition.Left
        txtNormalPositionRight.Text = .rcNormalPosition.Right
        txtNormalPositionTop.Text = .rcNormalPosition.Top
        txtNormalPositionBottom.Text = .rcNormalPosition.Bottom
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lvwWindow_ItemClick")
Resume Next
End Sub
