VERSION 5.00
Begin VB.Form frmFileInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Info"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmFileInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAdler32 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtLinks 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtIndexHi 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtIndexLo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtEncryptionStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ComboBox cboRound 
      Height          =   315
      Left            =   4320
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox cboOutput 
      Height          =   315
      Left            =   4320
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtCRC32 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdChecksum 
      Caption         =   "Checksum"
      Height          =   350
      Left            =   5040
      TabIndex        =   24
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
      Width           =   5895
   End
   Begin VB.DirListBox dirFileAttributes 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.FileListBox fileFileAttributes 
      Height          =   2040
      Hidden          =   -1  'True
      Left            =   120
      System          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.DriveListBox drvFileAttributes 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label lblEncryptionStatus 
      Caption         =   "Encryption Status"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblRound 
      Caption         =   "Round"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblSize 
      Caption         =   "Size"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblIndexHi 
      Caption         =   "Index Hi"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblIndexLo 
      Caption         =   "Index Lo"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblLinks 
      Caption         =   "Links"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblChecksum 
      Caption         =   "Checksum"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblCRC32 
      Caption         =   "CRC32"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblAdler32 
      Caption         =   "Adler32"
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   4320
      Width           =   1575
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
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmFileInfo"

Private Sub cboOutput_Click()
On Error GoTo VB_Error

    If txtSelected.Text <> vbNullString Then
        fileFileAttributes_Click
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboOutput_Click")
Resume Next
End Sub

Private Sub cboRound_Click()
On Error GoTo VB_Error

    If txtSelected.Text <> vbNullString Then
        fileFileAttributes_Click
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cboRound_Click")
Resume Next
End Sub

Private Sub cmdChecksum_Click()
On Error GoTo VB_Error

    If txtSelected.Text <> vbNullString Then
        If File_Size_Name(txtSelected.Text) < 2147483648# Then
            Dim sFileContents As String
            sFileContents = File_Read_Name(txtSelected.Text, CLng(File_Size_Name(txtSelected.Text)), 0)
            
            Dim crc As Long
            Dim adler As Long
            
            crc = crc32(crc, sFileContents, Len(sFileContents))
            adler = adler32(adler, sFileContents, Len(sFileContents))
            
            txtCRC32.Text = Right$("00000000" & ltoa_(crc, 16), 8)
            txtAdler32.Text = Right$("00000000" & ltoa_(adler, 16), 8)
        Else
            If MessageBoxEx(0&, "File is larger than 2GB.", "Restart", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdChecksum_Click", "MessageBoxEx")
        End If
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdChecksum_Click")
Resume Next
End Sub

Private Sub dirFileAttributes_Change()
On Error GoTo VB_Error

    fileFileAttributes.Path = dirFileAttributes.Path
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\dirFileAttributes_Change")
Resume Next
End Sub

Private Sub drvFileAttributes_Change()
On Error GoTo VB_Error
    
    dirFileAttributes.Path = drvFileAttributes.Drive
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\drvFileAttributes_Change")
Resume Next
End Sub

Private Sub fileFileAttributes_Click()
On Error GoTo VB_Error

    txtSelected.Text = Str_BckSlhTerm_Fix(dirFileAttributes.Path) & "\" & fileFileAttributes.FileName
    
    Dim BY_HANDLE_FILE_INFORMATION As BY_HANDLE_FILE_INFORMATION
    Dim hFile As Long
    
    hFile = CreateFile(txtSelected.Text, 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&): If hFile = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\fileFileAttributes_Click", "CreateFile")
    If GetFileInformationByHandle(hFile, BY_HANDLE_FILE_INFORMATION) = False Then Call Error_API(Err.LastDllError, sLocation & "\fileFileAttributes_Click", "GetFileInformationByHandle")
    If CloseHandle(hFile) = False Then Call Error_API(Err.LastDllError, sLocation & "\fileFileAttributes_Click", "CloseHandle")
    
    With BY_HANDLE_FILE_INFORMATION
        txtSize.Text = FormatNumber$(int32x32_int64(.nFileSizeLow, .nFileSizeHigh) / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex, , , True)
        txtIndexLo.Text = int32_uint32(.nFileIndexLow)
        txtIndexHi.Text = int32_uint32(.nFileIndexHigh)
        txtLinks.Text = FormatNumber$(int32_uint32(.nNumberOfLinks), 0, , , True)
    End With
    
    
    If Function_Exist("advapi32.dll", "FileEncryptionStatusA") = True Then
        Dim lStatus As Long
        If FileEncryptionStatus(txtSelected.Text, lStatus) = False Then Call Error_API(Err.LastDllError, sLocation & "\fileFileAttributes_Click", "FileEncryptionStatus")
        
        Select Case lStatus
            Case FILE_ENCRYPTABLE: txtEncryptionStatus.Text = "Encryptable"
            Case FILE_IS_ENCRYPTED: txtEncryptionStatus.Text = "Encrypted"
            Case FILE_SYSTEM_ATTR: txtEncryptionStatus.Text = "N/A - System File"
            Case FILE_ROOT_DIR: txtEncryptionStatus.Text = "N/A - Root Directory"
            Case FILE_SYSTEM_DIR: txtEncryptionStatus.Text = "N/A - System Directory"
            Case FILE_UNKNOWN: txtEncryptionStatus.Text = "Unknown"
            Case FILE_SYSTEM_NOT_SUPPORT: txtEncryptionStatus.Text = "Not Supported"
            Case FILE_USER_DISALLOWED: txtEncryptionStatus.Text = "User Disallowed"
            Case FILE_READ_ONLY: txtEncryptionStatus.Text = "N/A - Read Only"
            Case Else: txtEncryptionStatus.Text = "Unknown " & lStatus
        End Select
    End If
    
    
    txtCRC32.Text = vbNullString
    txtAdler32.Text = vbNullString
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\fileFileAttributes_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    With cboOutput
        .AddItem "Bytes"
        .AddItem "Kilobytes"
        .AddItem "Megabytes"
        .AddItem "Gigabytes"
        .AddItem "Terabytes"
    End With
    With cboRound
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
    End With
    
    cboOutput.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Output"), 0, 4)
    cboRound.ListIndex = MinMax(Reg_Read(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Round"), 0, 5)
    
    If Function_Exist("advapi32.dll", "FileEncryptionStatusA") = False Then
        lblEncryptionStatus.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Output", cboOutput.ListIndex, REG_DWORD)
    Call Reg_Write(HKEY_CURRENT_USER, sRegKey & "\FileInfo", "Round", cboRound.ListIndex, REG_DWORD)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
