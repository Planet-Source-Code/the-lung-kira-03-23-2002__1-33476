VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PE"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmPE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox drvFileTime 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
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
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get Info"
      Height          =   350
      Left            =   6120
      TabIndex        =   6
      Top             =   4680
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwPE 
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
Attribute VB_Name = "frmPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmPE"


Private Sub cmdGetInfo_Click()
On Error GoTo VB_Error
    
    Call ListView_Clear(lvwPE)
    
    
    If txtSelected.Text = vbNullString Then Exit Sub
    
    Dim lFileSize As Long
    lFileSize = File_Size_Name(txtSelected.Text)
    If lFileSize = 0 Then Exit Sub
    
    
    Dim sFileContents As String
    Dim IMAGE_DOS_HEADER As IMAGE_DOS_HEADER
    Dim IMAGE_FILE_HEADER As IMAGE_FILE_HEADER
    
    sFileContents = File_Read_Name(txtSelected.Text, lFileSize, 0)
    If Len(sFileContents) < Len(IMAGE_DOS_HEADER) Then Exit Sub
    
    Call MoveMemory(IMAGE_DOS_HEADER, ByVal sFileContents, Len(IMAGE_DOS_HEADER))
    
    If IMAGE_DOS_HEADER.e_magic <> IMAGE_DOS_SIGNATURE Then
        If IMAGE_DOS_HEADER.e_magic <> &H5A4D Then Exit Sub
    End If
    
    
    If Len(sFileContents) < ((Len(IMAGE_DOS_HEADER) + 5) + IMAGE_SIZEOF_FILE_HEADER) Then Exit Sub
    
    Dim lSignature As Long
    Call MoveMemory(lSignature, ByVal Mid$(sFileContents, IMAGE_DOS_HEADER.e_lfanew + 1, 4), 4)
    If lSignature <> IMAGE_NT_SIGNATURE Then
        If lSignature <> &H4550 Then Exit Sub
    End If
    
    Call MoveMemory(IMAGE_FILE_HEADER, ByVal Mid$(sFileContents, IMAGE_DOS_HEADER.e_lfanew + 5, IMAGE_SIZEOF_FILE_HEADER), IMAGE_SIZEOF_FILE_HEADER)
    
    
    With IMAGE_FILE_HEADER
        lvwPE.ListItems.Add(, , "Header Offset").SubItems(1) = FormatNumber(int32_uint32(IMAGE_DOS_HEADER.e_lfanew), 0, , , True)
        lvwPE.ListItems.Add(, , vbNullString).SubItems(1) = vbNullString
        
        lvwPE.ListItems.Add(, , "Signature").SubItems(1) = CBool(1)
        lvwPE.ListItems.Add(, , vbNullString).SubItems(1) = vbNullString
        
        lvwPE.ListItems.Add(, , "Image File Header").SubItems(1) = vbNullString
        Dim sData As String
        Select Case IMAGE_FILE_HEADER.Machine
            Case IMAGE_FILE_MACHINE_UNKNOWN: sData = "Unknown"
            Case IMAGE_FILE_MACHINE_I386: sData = "Intel 386"
            Case &H160: sData = "MIPS R3000 big-endian"
            Case IMAGE_FILE_MACHINE_R3000: sData = "MIPS R3000 little-endian"
            Case IMAGE_FILE_MACHINE_R4000: sData = "MIPS R4000 little-endian"
            Case IMAGE_FILE_MACHINE_R10000: sData = "MIPS R10000 little-endian"
            Case IMAGE_FILE_MACHINE_WCEMIPSV2: sData = "MIPS little-endian WCE v2"
            Case IMAGE_FILE_MACHINE_ALPHA: sData = "Alpha_AXP"
            Case IMAGE_FILE_MACHINE_SH3: sData = "SH3 little-endian"
            Case IMAGE_FILE_MACHINE_SH3DSP: sData = "SH3DSP"
            Case IMAGE_FILE_MACHINE_SH3E: sData = "SH3E little-endian"
            Case IMAGE_FILE_MACHINE_SH4: sData = "SH4 little-endian"
            Case IMAGE_FILE_MACHINE_SH5: sData = "SH5"
            Case IMAGE_FILE_MACHINE_ARM: sData = "ARM Little-Endian"
            Case IMAGE_FILE_MACHINE_THUMB: sData = "THUMB"
            Case IMAGE_FILE_MACHINE_AM33: sData = "AM33"
            Case IMAGE_FILE_MACHINE_POWERPC: sData = "IBM PowerPC Little-Endian"
            Case IMAGE_FILE_MACHINE_POWERPCFP: sData = "POWERPCFP"
            Case IMAGE_FILE_MACHINE_IA64: sData = "Intel 64"
            Case IMAGE_FILE_MACHINE_MIPS16: sData = "MIPS16"
            Case IMAGE_FILE_MACHINE_ALPHA64: sData = "ALPHA64"
            Case IMAGE_FILE_MACHINE_MIPSFPU: sData = "MIPSFPU"
            Case IMAGE_FILE_MACHINE_MIPSFPU16: sData = "MIPSFPU16"
            Case IMAGE_FILE_MACHINE_TRICORE: sData = "Infineon"
            Case IMAGE_FILE_MACHINE_CEF: sData = "CEF"
            Case IMAGE_FILE_MACHINE_EBC: sData = "EFI Byte Code"
            Case IMAGE_FILE_MACHINE_AMD64: sData = "AMD64 (K8)"
            Case IMAGE_FILE_MACHINE_M32R: sData = "M32R little-endian"
            Case IMAGE_FILE_MACHINE_CEE: sData = "CEE"
            Case Else: sData = "Unknown " & IMAGE_FILE_HEADER.Machine
        End Select
        lvwPE.ListItems.Add(, , "Machine").SubItems(1) = sData
        
        lvwPE.ListItems.Add(, , "Number Of Sections").SubItems(1) = FormatNumber(int16_uint16(IMAGE_FILE_HEADER.NumberOfSections), 0, , , True)
        lvwPE.ListItems.Add(, , "Time Date Stamp").SubItems(1) = DateAdd("s", int32_uint32(IMAGE_FILE_HEADER.TimeDateStamp), "12/31/1969 4:00:00 PM")
        lvwPE.ListItems.Add(, , "Pointer To Symbol Table").SubItems(1) = FormatNumber(int32_uint32(IMAGE_FILE_HEADER.PointerToSymbolTable), 0, , , True)
        lvwPE.ListItems.Add(, , "Number Of Symbols").SubItems(1) = FormatNumber(int32_uint32(IMAGE_FILE_HEADER.NumberOfSymbols), 0, , , True)
        lvwPE.ListItems.Add(, , "Size Of Optional Header").SubItems(1) = FormatNumber(int16_uint16(IMAGE_FILE_HEADER.SizeOfOptionalHeader), 0, , , True)
        
        lvwPE.ListItems.Add(, , vbNullString).SubItems(1) = vbNullString
        lvwPE.ListItems.Add(, , "Characteristics").SubItems(1) = vbNullString
        lvwPE.ListItems.Add(, , "Relocation Info Stripped").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_RELOCS_STRIPPED)
        lvwPE.ListItems.Add(, , "Executable").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_EXECUTABLE_IMAGE)
        lvwPE.ListItems.Add(, , "Line Numbers Stripped").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_LINE_NUMS_STRIPPED)
        lvwPE.ListItems.Add(, , "Local Symbols Stripped").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_LOCAL_SYMS_STRIPPED)
        lvwPE.ListItems.Add(, , "Aggressive Trim Working Set").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_AGGRESIVE_WS_TRIM)
        lvwPE.ListItems.Add(, , "Large Address Aware").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_LARGE_ADDRESS_AWARE)
        lvwPE.ListItems.Add(, , "Low Machine Word Bytes Reserved").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_BYTES_REVERSED_LO)
        lvwPE.ListItems.Add(, , "32Bit Word Machine").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_32BIT_MACHINE)
        lvwPE.ListItems.Add(, , "Debugging Info Stripped").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_DEBUG_STRIPPED)
        lvwPE.ListItems.Add(, , "If Removable Run From Swap").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP)
        lvwPE.ListItems.Add(, , "If Net Run From Swap").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_NET_RUN_FROM_SWAP)
        lvwPE.ListItems.Add(, , "System File").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_SYSTEM)
        lvwPE.ListItems.Add(, , "DLL File").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_DLL)
        lvwPE.ListItems.Add(, , "Run On Up System Only").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_UP_SYSTEM_ONLY)
        lvwPE.ListItems.Add(, , "High Machine Word Bytes Reserved").SubItems(1) = CBool(IMAGE_FILE_HEADER.Characteristics And IMAGE_FILE_BYTES_REVERSED_HI)
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
    
    With lvwPE.ColumnHeaders
        .Add , , "Identifier"
        .Add , , "Value"
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
