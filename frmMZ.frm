VERSION 5.00
Begin VB.Form frmMZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MZ"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmMZ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
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
   Begin VB.TextBox txtOverlay 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtRelocationOffset 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtInitialCS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtInitialIP 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtChecksum 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtInitialSP 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtInitialSS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtMaxPara 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtMinPara 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txt16Para 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtRelocationTables 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txt512Pages 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtSizeMod 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox chkSignature 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get Info"
      Height          =   350
      Left            =   6120
      TabIndex        =   33
      Top             =   4440
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
   Begin VB.Label lblOverlay 
      Caption         =   "Overlay Number"
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label lblRelocationOffset 
      Caption         =   "Relocation Offset"
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label lblInitialIP 
      Caption         =   "Initial IP Value"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label lblInitialCS 
      Caption         =   "Initial Relative CS Value (Paragraphs)"
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label lblChecksum 
      Caption         =   "Checksum"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label lblInitialSP 
      Caption         =   "Initial SP Value"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label lblInitialSS 
      Caption         =   "Initial Relative SS Value (Paragraphs)"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label lblMinPara 
      Caption         =   "Minimum Number of Paragraphs"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblMaxPara 
      Caption         =   "Maximum Number of Paragraphs"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label lbl16Para 
      Caption         =   "16b Paragraphs for Header/Relocation Table"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblRelocationTables 
      Caption         =   "Relocation Tables"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lbl512Pages 
      Caption         =   "512b Pages"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblSizeMod 
      Caption         =   "Image Size Mod 512"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblSignature 
      Caption         =   "Signature"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "frmMZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMZ"


Private Sub cmdGetInfo_Click()
On Error GoTo VB_Error

    chkSignature.value = 0
    txtSizeMod.Text = vbNullString
    txt512Pages.Text = vbNullString
    txtRelocationTables.Text = vbNullString
    txt16Para.Text = vbNullString
    txtMinPara.Text = vbNullString
    txtMaxPara.Text = vbNullString
    txtInitialSS.Text = vbNullString
    txtInitialSP.Text = vbNullString
    txtChecksum.Text = vbNullString
    txtInitialIP.Text = vbNullString
    txtInitialCS.Text = vbNullString
    txtRelocationOffset.Text = vbNullString
    txtOverlay.Text = vbNullString
    
    
    If txtSelected.Text = vbNullString Then Exit Sub
    
    If File_Size_Name(txtSelected.Text) < 28 Then Exit Sub
    
    
    Dim sFileContents As String
    sFileContents = File_Read_Name(txtSelected.Text, 28, 0)
    If Len(sFileContents) < 28 Then Exit Sub
    
    Dim IMAGE_DOS_HEADER As IMAGE_DOS_HEADER
    Call MoveMemory(IMAGE_DOS_HEADER, ByVal sFileContents, 28)
    
    
    If IMAGE_DOS_HEADER.e_magic <> IMAGE_DOS_SIGNATURE Then
        If IMAGE_DOS_HEADER.e_magic <> &H5A4D Then Exit Sub
    End If
    
    chkSignature.value = 1
    With IMAGE_DOS_HEADER
        txtSizeMod.Text = int16_uint16(.e_cblp)
        txt512Pages.Text = int16_uint16(.e_cp)
        txtRelocationTables.Text = int16_uint16(.e_crlc)
        txt16Para.Text = int16_uint16(.e_cparhdr)
        txtMinPara.Text = int16_uint16(.e_minalloc)
        txtMaxPara.Text = int16_uint16(.e_maxalloc)
        txtInitialSS.Text = int16_uint16(.e_ss)
        txtInitialSP.Text = int16_uint16(.e_sp)
        txtChecksum.Text = int16_uint16(.e_csum)
        txtInitialIP.Text = int16_uint16(.e_ip)
        txtInitialCS.Text = int16_uint16(.e_cs)
        txtRelocationOffset.Text = int16_uint16(.e_lfarlc)
        txtOverlay.Text = int16_uint16(.e_ovno)
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
