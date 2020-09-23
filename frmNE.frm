VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NE"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmNE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get Info"
      Height          =   350
      Left            =   6120
      TabIndex        =   6
      Top             =   4680
      Width           =   975
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
   Begin VB.FileListBox fileFileTime 
      Height          =   2040
      Hidden          =   -1  'True
      Left            =   120
      System          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.DirListBox dirFileTime 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.DriveListBox drvFileTime 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwNE 
      Height          =   3735
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6588
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
   Begin VB.Label lblSelected 
      Caption         =   "Selected"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmNE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmNE"


Private Sub cmdGetInfo_Click()
On Error GoTo VB_Error
    
    Call ListView_Clear(lvwNE)
    
    
    If txtSelected.Text = vbNullString Then Exit Sub
    
    Dim lFileSize As Long
    lFileSize = File_Size_Name(txtSelected.Text)
    If lFileSize = 0 Then Exit Sub
    
    
    Dim sFileContents As String
    Dim IMAGE_DOS_HEADER As IMAGE_DOS_HEADER
    Dim IMAGE_OS2_HEADER As IMAGE_OS2_HEADER
    
    sFileContents = File_Read_Name(txtSelected.Text, lFileSize, 0)
    If Len(sFileContents) < Len(IMAGE_DOS_HEADER) Then Exit Sub
    
    Call MoveMemory(IMAGE_DOS_HEADER, ByVal sFileContents, Len(IMAGE_DOS_HEADER))
    
    
    If IMAGE_DOS_HEADER.e_magic <> IMAGE_DOS_SIGNATURE Then
        If IMAGE_DOS_HEADER.e_magic <> &H5A4D Then Exit Sub
    End If
    
    
    If Len(sFileContents) < ((Len(IMAGE_DOS_HEADER) + 1) + Len(IMAGE_OS2_HEADER)) Then Exit Sub
    
    Call MoveMemory(IMAGE_OS2_HEADER, ByVal Mid$(sFileContents, IMAGE_DOS_HEADER.e_lfanew + 1, Len(IMAGE_OS2_HEADER)), Len(IMAGE_OS2_HEADER))
    
    
    If IMAGE_OS2_HEADER.ne_magic <> IMAGE_OS2_SIGNATURE Then
        If IMAGE_OS2_HEADER.ne_magic <> &H454E Then Exit Sub
    End If
    
    
    With IMAGE_OS2_HEADER
        lvwNE.ListItems.Add(, , "Header Offset").SubItems(1) = FormatNumber(int32_uint32(IMAGE_DOS_HEADER.e_lfanew), 0, , , True)
        lvwNE.ListItems.Add(, , vbNullString).SubItems(1) = vbNullString
        
        lvwNE.ListItems.Add(, , "Signature").SubItems(1) = CBool(1)
        lvwNE.ListItems.Add(, , "Linker Version").SubItems(1) = .ne_ver
        lvwNE.ListItems.Add(, , "Linker Revision").SubItems(1) = .ne_rev
        lvwNE.ListItems.Add(, , "Entry Table Offset Relative").SubItems(1) = FormatNumber(int16_uint16(.ne_enttab), 0, , , True)
        lvwNE.ListItems.Add(, , "Bytes In Entry Table").SubItems(1) = FormatNumber(int16_uint16(.ne_cbenttab), 0, , , True)
        lvwNE.ListItems.Add(, , "Entire File CRC").SubItems(1) = FormatNumber(int16_uint16(.ne_crc), 0, , , True)
        
        lvwNE.ListItems.Add(, , vbNullString).SubItems(1) = vbNullString
        lvwNE.ListItems.Add(, , "Flag Word").SubItems(1) = vbNullString
        
        Dim sData As String
        Select Case True
            Case .ne_flags And &H0: sData = "NOAUTODATA"
            Case .ne_flags And &H1: sData = "SINGLEDATA"
            Case .ne_flags And &H2: sData = "MULTIPLEDATA"
            Case Else: sData = "Unknown " & .ne_flags
        End Select
        lvwNE.ListItems.Add(, , "Data Type").SubItems(1) = sData
        
        lvwNE.ListItems.Add(, , "First Segment Is Code").SubItems(1) = CBool(.ne_flags And &H800)
        lvwNE.ListItems.Add(, , "Link Time Error").SubItems(1) = CBool(.ne_flags And &H2000)
        lvwNE.ListItems.Add(, , "Library Module").SubItems(1) = CBool(.ne_flags And &H8000)
        
        lvwNE.ListItems.Add(, , vbNullString).SubItems(1) = vbNullString
        
        
        lvwNE.ListItems.Add(, , "Automatic Data Segment Number").SubItems(1) = FormatNumber(int16_uint16(.ne_autodata), 0, , , True)
        lvwNE.ListItems.Add(, , "Local Heap Size Bytes").SubItems(1) = FormatNumber(int16_uint16(.ne_heap), 0, , , True)
        lvwNE.ListItems.Add(, , "Local Stack Size Bytes").SubItems(1) = FormatNumber(int16_uint16(.ne_stack), 0, , , True)
        lvwNE.ListItems.Add(, , "Initial CS:IP").SubItems(1) = FormatNumber(int32_uint32(.ne_csip), 0, , , True)
        lvwNE.ListItems.Add(, , "Initial SS:SP").SubItems(1) = FormatNumber(int32_uint32(.ne_sssp), 0, , , True)
        lvwNE.ListItems.Add(, , "Segment Table Entries").SubItems(1) = FormatNumber(int16_uint16(.ne_cseg), 0, , , True)
        lvwNE.ListItems.Add(, , "Module Reference Table Entries").SubItems(1) = FormatNumber(int16_uint16(.ne_cmod), 0, , , True)
        lvwNE.ListItems.Add(, , "Non-Resident Name Table Bytes").SubItems(1) = FormatNumber(int16_uint16(.ne_cbnrestab), 0, , , True)
        lvwNE.ListItems.Add(, , "Segment Table Offset Relative").SubItems(1) = FormatNumber(int16_uint16(.ne_segtab), 0, , , True)
        lvwNE.ListItems.Add(, , "Resource Table Offset Relative").SubItems(1) = FormatNumber(int16_uint16(.ne_rsrctab), 0, , , True)
        lvwNE.ListItems.Add(, , "Resident Name Table Offset Relative").SubItems(1) = FormatNumber(int16_uint16(.ne_restab), 0, , , True)
        lvwNE.ListItems.Add(, , "Module Reference Table Offset Relative").SubItems(1) = FormatNumber(int16_uint16(.ne_modtab), 0, , , True)
        lvwNE.ListItems.Add(, , "Imported Name Table Offset Relative").SubItems(1) = FormatNumber(int16_uint16(.ne_imptab), 0, , , True)
        lvwNE.ListItems.Add(, , "Non-Resident Name Table Offset Relative").SubItems(1) = FormatNumber(int32_uint32(.ne_nrestab), 0, , , True)
        lvwNE.ListItems.Add(, , "Entry Table Movable Entries").SubItems(1) = FormatNumber(int16_uint16(.ne_cmovent), 0, , , True)
        lvwNE.ListItems.Add(, , "Segment Alignment Shift Count").SubItems(1) = FormatNumber(int16_uint16(.ne_align), 0, , , True)
        lvwNE.ListItems.Add(, , "Resource Segment Count").SubItems(1) = FormatNumber(int16_uint16(.ne_cres), 0, , , True)
        
        Select Case True
            Case .ne_exetyp And &H2: sData = "Windows"
            Case Else: sData = "Unknown " & .ne_exetyp
        End Select
        lvwNE.ListItems.Add(, , "Target Operating System").SubItems(1) = sData
        
        
        lvwNE.ListItems.Add(, , vbNullString).SubItems(1) = vbNullString
        lvwNE.ListItems.Add(, , "Other EXE Flags").SubItems(1) = vbNullString
        
        lvwNE.ListItems.Add(, , "Win2.x App Run In Win3.x PMode").SubItems(1) = CBool(.ne_flagsothers And &H2)
        lvwNE.ListItems.Add(, , "Win2.x App Supports Proportional Fonts").SubItems(1) = CBool(.ne_flagsothers And &H4)
        lvwNE.ListItems.Add(, , "Fast Load Area").SubItems(1) = CBool(.ne_flagsothers And &H8)
        
        lvwNE.ListItems.Add(, , vbNullString).SubItems(1) = vbNullString
        
        
        lvwNE.ListItems.Add(, , "Return Thunk Offset Relative").SubItems(1) = FormatNumber(int16_uint16(.ne_pretthunks), 0, , , True)
        lvwNE.ListItems.Add(, , "Segment Reference Bytes Offset Relative").SubItems(1) = FormatNumber(int16_uint16(.ne_psegrefbytes), 0, , , True)
        lvwNE.ListItems.Add(, , "Minimum Code Swap Area Size").SubItems(1) = FormatNumber(int16_uint16(.ne_swaparea), 0, , , True)
        lvwNE.ListItems.Add(, , "Expected Windows Version").SubItems(1) = FormatNumber(int16_uint16(.ne_expver), 0, , , True)
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdGetInfo_Click")
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
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\fileFileTime_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    With lvwNE.ColumnHeaders
        .Add , , "Identifier"
        .Add , , "Value"
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
