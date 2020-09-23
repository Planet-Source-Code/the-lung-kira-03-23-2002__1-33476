VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWindowSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Settings"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmWindowSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstThread 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CheckBox chkShowWindow 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   38
      Top             =   4440
      Width           =   255
   End
   Begin VB.CheckBox chkNoZOrder 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   36
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox chkNoSize 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   34
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chkNoSendChanging 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   32
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkNoRedraw 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   30
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkNoOwnerZOrder 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkNoMove 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkNoCopyBits 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   24
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkNoActivate 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkHideWindow 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkFrameChanged 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkDeferErase 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkAsyncWindowPos 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
   Begin VB.ComboBox cboZOrder 
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
      Left            =   4800
      TabIndex        =   12
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CheckBox chkInvertFlash 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtWindowText 
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5760
      TabIndex        =   39
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2040
      TabIndex        =   6
      Top             =   5520
      Width           =   975
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
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
   Begin VB.Label lblThread 
      Caption         =   "Thread"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
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
   Begin VB.Label lblWindow 
      Caption         =   "Window"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblShowWindow 
      Caption         =   "Show Window"
      Height          =   255
      Left            =   3240
      TabIndex        =   37
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblNoZOrder 
      Caption         =   "No Z Order"
      Height          =   255
      Left            =   3240
      TabIndex        =   35
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblNoSize 
      Caption         =   "No Size"
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblNoSendChanging 
      Caption         =   "No Send Changing"
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblNoRedraw 
      Caption         =   "No Redraw"
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblNoOwnerZOrder 
      Caption         =   "No Owner Z Order"
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblNoMove 
      Caption         =   "No Move"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblNoCopyBits 
      Caption         =   "No Copy Bits"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblNoActivate 
      Caption         =   "No Activate"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblHideWindow 
      Caption         =   "Hide Window"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblFrameChanged 
      Caption         =   "Frame Changed"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblDeferErase 
      Caption         =   "Defer Erase"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblAsyncWindowPos 
      Caption         =   "Async Window Pos"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblZOrder 
      Caption         =   "Z Order"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblInvertFlash 
      Caption         =   "Invert Flash"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblWindow_Text 
      Caption         =   "Window Text"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmWindowSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmWindowSettings"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    If lvwWindow.SelectedItem Is Nothing Then Exit Sub
    
    
    Dim lHandle As Long
    lHandle = lvwWindow.SelectedItem
    
    If SetWindowText(lHandle, txtWindowText.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetWindowText")
    FlashWindow lHandle, CBool(chkInvertFlash.value)
    
    
    Dim lInsert As Long
    Dim lFlags As Long
    
    Select Case cboZOrder.ListIndex
        Case 0: lInsert = HWND_BOTTOM
        Case 1: lInsert = HWND_NOTOPMOST
        Case 2: lInsert = HWND_TOP
        Case 3: lInsert = HWND_TOPMOST
    End Select
    
    If chkAsyncWindowPos.value = 1 Then lFlags = lFlags Or SWP_ASYNCWINDOWPOS
    If chkDeferErase.value = 1 Then lFlags = lFlags Or SWP_DEFERERASE
    If chkFrameChanged.value = 1 Then lFlags = lFlags Or SWP_FRAMECHANGED
    If chkHideWindow.value = 1 Then lFlags = lFlags Or SWP_HIDEWINDOW
    If chkNoActivate.value = 1 Then lFlags = lFlags Or SWP_NOACTIVATE
    If chkNoCopyBits.value = 1 Then lFlags = lFlags Or SWP_NOCOPYBITS
    If chkNoMove.value = 1 Then lFlags = lFlags Or SWP_NOMOVE
    If chkNoOwnerZOrder.value = 1 Then lFlags = lFlags Or SWP_NOOWNERZORDER
    If chkNoRedraw.value = 1 Then lFlags = lFlags Or SWP_NOREDRAW
    If chkNoSendChanging.value = 1 Then lFlags = lFlags Or SWP_NOSENDCHANGING
    If chkNoSize.value = 1 Then lFlags = lFlags Or SWP_NOSIZE
    If chkNoZOrder.value = 1 Then lFlags = lFlags Or SWP_NOZORDER
    If chkShowWindow.value = 1 Then lFlags = lFlags Or SWP_SHOWWINDOW
    
    
    Dim WINDOWPLACEMENT As WINDOWPLACEMENT
    WINDOWPLACEMENT.Length = Len(WINDOWPLACEMENT)
    If GetWindowPlacement(lHandle, WINDOWPLACEMENT) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "GetWindowPlacement")
    
    With WINDOWPLACEMENT.rcNormalPosition
        If SetWindowPos(lHandle, lInsert, .Left, .Top, .Right - .Left, .Bottom - .Top, lFlags) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetWindowPos")
    End With
    
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
    
    
    txtWindowText.Text = vbNullString
    
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
    
    With cboZOrder
        .AddItem "Bottom"
        .AddItem "Not Top Most"
        .AddItem "Top"
        .AddItem "Top Most"
        .ListIndex = 1
    End With
    
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub lstThread_Click()
On Error GoTo VB_Error
    
    Call ListView_Clear(lvwWindow)
    
    If EnumWindows(AddressOf frmWindowSettings_EnumWindowsProc, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "EnumWindows")
    
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
    
    
    txtWindowText.Text = lvwWindow.SelectedItem.SubItems(1)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lvwWindow_ItemClick")
Resume Next
End Sub
