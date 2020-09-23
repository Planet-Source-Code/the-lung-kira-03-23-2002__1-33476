VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHeaps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Heaps"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmHeaps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtListSize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtB2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "B"
      Top             =   3240
      Width           =   135
   End
   Begin VB.TextBox txtB3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "B"
      Top             =   4920
      Width           =   135
   End
   Begin VB.TextBox txtLockCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtHeapHandle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtFlags 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtBlockSize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CheckBox chkDefaultHeap 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   2880
      Width           =   255
   End
   Begin VB.ListBox lstHeapList 
      Height          =   1620
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2895
   End
   Begin VB.ListBox lstHeap 
      Height          =   1620
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4920
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
   Begin VB.Label lblListSize 
      Caption         =   "List Size"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblHeapHandle 
      Caption         =   "Heap Handle"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lblBlockSize 
      Caption         =   "Block Size"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblLockCount 
      Caption         =   "Lock Count"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblFlags 
      Caption         =   "Flags"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblDefaultHeap 
      Caption         =   "Default Heap"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblHeapList 
      Caption         =   "Heap List ID"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblHeap 
      Caption         =   "Heap Address"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4680
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
Attribute VB_Name = "frmHeaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Process() As PROCESSENTRY32
Dim HeapList() As HEAPLIST32
Dim Heap() As HEAPENTRY32
Const sLocation As String = "frmHeaps"


Private Sub cmdRefresh_Click()
On Error GoTo VB_Error

    Call ListView_Clear(lvwProcess)
    lstHeapList.Clear
    lstHeap.Clear
    
    chkDefaultHeap.value = 0
    txtListSize.Text = vbNullString
    txtBlockSize.Text = vbNullString
    txtFlags.Text = vbNullString
    txtHeapHandle.Text = vbNullString
    txtLockCount.Text = vbNullString
    
    
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
        lblHeapList.Enabled = False
        lstHeapList.Enabled = False
        lblHeap.Enabled = False
        lstHeap.Enabled = False
        lblDefaultHeap.Enabled = False
        lblListSize.Enabled = False
        lblBlockSize.Enabled = False
        lblFlags.Enabled = False
        lblHeapHandle.Enabled = False
        lblLockCount.Enabled = False
        cmdRefresh.Enabled = False
    Else
        Call cmdRefresh_Click
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub lstHeap_Click()
On Error GoTo VB_Error

    With Heap(lstHeap.ListIndex)
        txtBlockSize.Text = FormatNumber(int32_uint32(.dwBlockSize), 0, , , True)
        
        Select Case .dwFlags
            Case LF32_FIXED: txtFlags.Text = "Fixed"
            Case LF32_FREE: txtFlags.Text = "Free"
            Case LF32_MOVEABLE: txtFlags.Text = "Moveable"
            Case Else: txtFlags.Text = int32_uint32(.dwFlags)
        End Select
        
        txtHeapHandle.Text = int32_uint32(.hHandle)
        txtLockCount.Text = FormatNumber(int32_uint32(.dwLockCount), 0, , , True)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstHeap_Click")
Resume Next
End Sub

Private Sub lstHeapList_Click()
On Error GoTo VB_Error
    
    If lvwProcess.SelectedItem Is Nothing Then Exit Sub
    
    
    lstHeap.Clear
    
    txtBlockSize.Text = vbNullString
    txtFlags.Text = vbNullString
    txtHeapHandle.Text = vbNullString
    txtLockCount.Text = vbNullString
    
    
    Dim lCount As Long
    
    lCount = Heap32_Enum(Heap(), lvwProcess.SelectedItem, lstHeapList.List(lstHeapList.ListIndex))
    
    Dim lIncrement As Long
    Dim lSize As Long
    For lIncrement = 0 To lCount
        lstHeap.AddItem Heap(lIncrement).dwAddress
        lSize = lSize + Heap(lIncrement).dwBlockSize
    Next lIncrement
    
    chkDefaultHeap.value = IIf(HeapList(lstHeapList.ListIndex).dwFlags And HF32_DEFAULT, 1, 0)
    txtListSize.Text = FormatNumber(lSize, 0, , , True)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstHeapList_Click")
Resume Next
End Sub

Private Sub lvwProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo VB_Error
    
    If lvwProcess.SelectedItem Is Nothing Then Exit Sub
    
    
    lstHeapList.Clear
    lstHeap.Clear
    
    chkDefaultHeap.value = 0
    txtListSize.Text = vbNullString
    txtBlockSize.Text = vbNullString
    txtFlags.Text = vbNullString
    txtHeapHandle.Text = vbNullString
    txtLockCount.Text = vbNullString
    
    
    Dim lCount As Long
    
    lCount = Heap32List_Enum(HeapList(), lvwProcess.SelectedItem)
    
    Dim lIncrement As Long
    Dim lSize As Long
    For lIncrement = 0 To lCount
        lstHeapList.AddItem HeapList(lIncrement).th32HeapID
    Next lIncrement
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lvwProcess_ItemClick")
Resume Next
End Sub
