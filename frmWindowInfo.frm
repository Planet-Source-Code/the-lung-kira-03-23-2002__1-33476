VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWindowInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Info"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmWindowInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstThread 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtPIX2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtPIX1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "PIX"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtAssociatedModule 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtClientCoordinatesBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox txtClientCoordinatesTop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtClientCoordinatesRight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox txtZOrderPrevious 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtZOrderNext 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CheckBox chkUnicode 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   6240
      TabIndex        =   30
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtRootOwner 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtRoot 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtParent 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtOwner 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtCreatorVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtClassAtom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtWindowBorderHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox lstExtendedStyles 
      Height          =   645
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   47
      Top             =   4200
      Width           =   1575
   End
   Begin VB.ListBox lstStyles 
      Height          =   645
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   36
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtWindowBorderWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtClientCoordinatesLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtWindowText 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   1575
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
   Begin VB.Label lblAssociatedModule 
      Caption         =   "Associated Module"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblUnicode 
      Caption         =   "Unicode"
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblZOrderPrevious 
      Caption         =   "Z Order Previous"
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblZOrderNext 
      Caption         =   "Z Order Next"
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblOwner 
      Caption         =   "Owner"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblParent 
      Caption         =   "Parent"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblRoot 
      Caption         =   "Root"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblRootOwner 
      Caption         =   "Root Owner"
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblExtendedStyles 
      Caption         =   "Extended Styles"
      Height          =   255
      Left            =   3240
      TabIndex        =   46
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblStyles 
      Caption         =   "Styles"
      Height          =   255
      Left            =   3240
      TabIndex        =   35
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblCreatorVersion 
      Caption         =   "Creator Version"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblClassAtom 
      Caption         =   "Class Atom"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblWindowBorderHeight 
      Caption         =   "Border Height"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblWindowBorderWidth 
      Caption         =   "Border Width"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinatesBottom 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   3240
      TabIndex        =   44
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinatesTop 
      Caption         =   "Top"
      Height          =   255
      Left            =   3240
      TabIndex        =   42
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinatesRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   3240
      TabIndex        =   40
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinates 
      Caption         =   "Client Coordinates"
      Height          =   255
      Left            =   3240
      TabIndex        =   37
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinatesLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   3240
      TabIndex        =   38
      Top             =   5160
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
Attribute VB_Name = "frmWindowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmWindowInfo"


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

    If Function_Exist("user32.dll", "GetAncestor") = False Then
        lblParent.Enabled = False
        lblRoot.Enabled = False
        lblRootOwner.Enabled = False
    End If
    If Function_Exist("user32.dll", "GetWindowInfo") = False Then
        lblClassAtom.Enabled = False
        lblClientCoordinates.Enabled = False
        lblClientCoordinatesLeft.Enabled = False
        lblClientCoordinatesRight.Enabled = False
        lblClientCoordinatesTop.Enabled = False
        lblClientCoordinatesBottom.Enabled = False
        lblCreatorVersion.Enabled = False
        lblExtendedStyles.Enabled = False
        lstExtendedStyles.Enabled = False
        lblStyles.Enabled = False
        lstStyles.Enabled = False
        lblWindowBorderWidth.Enabled = False
        lblWindowBorderHeight.Enabled = False
    End If
    If Function_Exist("user32.dll", "GetWindowModuleFileNameA") = False Then
        lblAssociatedModule.Enabled = False
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
    
    If EnumWindows(AddressOf frmWindowInfo_EnumWindowsProc, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdRefresh_Click", "EnumWindows")
    
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
    
    
    txtOwner.Text = GetWindow(lHandle, GW_OWNER)
    txtZOrderNext.Text = GetWindow(lHandle, GW_HWNDNEXT)
    txtZOrderPrevious.Text = GetWindow(lHandle, GW_HWNDPREV)
    chkUnicode.value = IsWindowUnicode(lHandle)
    
    If Function_Exist("user32.dll", "GetAncestor") = True Then
        txtParent.Text = GetAncestor(lHandle, GA_PARENT)
        txtRoot.Text = GetAncestor(lHandle, GA_ROOT)
        txtRootOwner.Text = GetAncestor(lHandle, GA_ROOTOWNER)
    End If
    If Function_Exist("user32.dll", "GetWindowInfo") = True Then
        Dim WINDOWINFO As WINDOWINFO
        WINDOWINFO.cbSize = Len(WINDOWINFO)
        If GetWindowInfo(lHandle, WINDOWINFO) = False Then Call Error_API(Err.LastDllError, sLocation & "\lvwWindow_ItemClick", "GetWindowInfo")
        
        With WINDOWINFO
            txtClientCoordinatesLeft.Text = .rcClient.Left
            txtClientCoordinatesRight.Text = .rcClient.Right
            txtClientCoordinatesTop.Text = .rcClient.Top
            txtClientCoordinatesBottom.Text = .rcClient.Bottom
            
            lstStyles.Clear
            If .dwStyle And WS_OVERLAPPED Then lstStyles.AddItem "Overlapped"
            If .dwStyle And WS_POPUP Then lstStyles.AddItem "Popup"
            If .dwStyle And WS_CHILD Then lstStyles.AddItem "Child"
            If .dwStyle And WS_MINIMIZE Then lstStyles.AddItem "Minimize"
            If .dwStyle And WS_VISIBLE Then lstStyles.AddItem "Visible"
            If .dwStyle And WS_DISABLED Then lstStyles.AddItem "Disabled"
            If .dwStyle And WS_CLIPSIBLINGS Then lstStyles.AddItem "Clip Siblings"
            If .dwStyle And WS_CLIPCHILDREN Then lstStyles.AddItem "Clib Children"
            If .dwStyle And WS_MAXIMIZE Then lstStyles.AddItem "Maximize"
            If .dwStyle And WS_CAPTION Then lstStyles.AddItem "Caption"
            If .dwStyle And WS_BORDER Then lstStyles.AddItem "Border"
            If .dwStyle And WS_DLGFRAME Then lstStyles.AddItem "Dialog Frame"
            If .dwStyle And WS_VSCROLL Then lstStyles.AddItem "Vertical Scroll"
            If .dwStyle And WS_HSCROLL Then lstStyles.AddItem "Horizontal Scroll"
            If .dwStyle And WS_SYSMENU Then lstStyles.AddItem "System Menu"
            If .dwStyle And WS_THICKFRAME Then lstStyles.AddItem "Thick Frame"
            If .dwStyle And WS_GROUP Then lstStyles.AddItem "Group"
            If .dwStyle And WS_TABSTOP Then lstStyles.AddItem "Tab Stop"
            If .dwStyle And WS_MINIMIZEBOX Then lstStyles.AddItem "Minimize Box"
            If .dwStyle And WS_MAXIMIZEBOX Then lstStyles.AddItem "Maximize Box"
            
            lstExtendedStyles.Clear
            If .dwExStyle And WS_EX_DLGMODALFRAME Then lstExtendedStyles.AddItem "Dialog Modal Frame"
            If .dwExStyle And WS_EX_NOPARENTNOTIFY Then lstExtendedStyles.AddItem "No Parent Notify"
            If .dwExStyle And WS_EX_TOPMOST Then lstExtendedStyles.AddItem "Top Most"
            If .dwExStyle And WS_EX_ACCEPTFILES Then lstExtendedStyles.AddItem "Accept Files"
            If .dwExStyle And WS_EX_TRANSPARENT Then lstExtendedStyles.AddItem "Transparent"
            If .dwExStyle And WS_EX_MDICHILD Then lstExtendedStyles.AddItem "MDI Child"
            If .dwExStyle And WS_EX_TOOLWINDOW Then lstExtendedStyles.AddItem "Tool Window"
            If .dwExStyle And WS_EX_WINDOWEDGE Then lstExtendedStyles.AddItem "Window Edge"
            If .dwExStyle And WS_EX_CLIENTEDGE Then lstExtendedStyles.AddItem "Client Edge"
            If .dwExStyle And WS_EX_CONTEXTHELP Then lstExtendedStyles.AddItem "Context Help"
            If .dwExStyle And WS_EX_RIGHT Then lstExtendedStyles.AddItem "Right"
            If .dwExStyle And WS_EX_LEFT Then lstExtendedStyles.AddItem "Left"
            If .dwExStyle And WS_EX_RTLREADING Then lstExtendedStyles.AddItem "Right Reading"
            If .dwExStyle And WS_EX_LEFTSCROLLBAR Then lstExtendedStyles.AddItem "Left Scroll Bar"
            If .dwExStyle And WS_EX_CONTROLPARENT Then lstExtendedStyles.AddItem "Control Parent"
            If .dwExStyle And WS_EX_STATICEDGE Then lstExtendedStyles.AddItem "Static Edge"
            If .dwExStyle And WS_EX_APPWINDOW Then lstExtendedStyles.AddItem "App Window"
            If .dwExStyle And WS_EX_LAYERED Then lstExtendedStyles.AddItem "Layered"
            If .dwExStyle And WS_EX_NOINHERITLAYOUT Then lstExtendedStyles.AddItem "No Inherit Layout"
            If .dwExStyle And WS_EX_LAYOUTRTL Then lstExtendedStyles.AddItem "Layout Right"
            If .dwExStyle And WS_EX_COMPOSITED Then lstExtendedStyles.AddItem "Composited"
            If .dwExStyle And WS_EX_NOACTIVATE Then lstExtendedStyles.AddItem "No Active"
            
            txtWindowBorderWidth.Text = int32_uint32(.cxWindowBorders)
            txtWindowBorderHeight.Text = int32_uint32(.cyWindowBorders)
            txtClassAtom.Text = .atomWindowType
            txtCreatorVersion.Text = .wCreatorVersion
        End With
    End If
    If Function_Exist("user32.dll", "GetWindowModuleFileNameA") = True Then
        Dim sFileName As String
        sFileName = String$(MAX_PATH, 0)
        txtAssociatedModule.Text = Left$(sFileName, GetWindowModuleFileName(lHandle, sFileName, Len(sFileName)))
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lvwWindow_ItemClick")
Resume Next
End Sub
