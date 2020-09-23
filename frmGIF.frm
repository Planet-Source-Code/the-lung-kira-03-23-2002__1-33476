VERSION 5.00
Begin VB.Form frmGIF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GIF"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmGIF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPIX2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      TabIndex        =   15
      Text            =   "PIX"
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtPIX1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      TabIndex        =   12
      Text            =   "PIX"
      Top             =   1680
      Width           =   255
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
   Begin VB.TextBox txtPixelAspectRatio 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtBackgroundColorIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CheckBox chkGlobalColorTableFlag 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   23
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txtColorResolution 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox chkSortFlag 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   19
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtGlobalColorTableSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtLogicalScreenHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtLogicalScreenWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get Info"
      Height          =   350
      Left            =   6120
      TabIndex        =   28
      Top             =   3840
      Width           =   975
   End
   Begin VB.CheckBox chkSignature 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   840
      Width           =   255
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
   Begin VB.Label lblLogicalScreenDescriptor 
      Caption         =   "Logical Screen Descriptor"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblBackgroundColorIndex 
      Caption         =   "Background Color Index"
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblPixelAspectRatio 
      Caption         =   "Pixel Aspect Ratio"
      Height          =   255
      Left            =   2520
      TabIndex        =   26
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblGlobalColorTableFlag 
      Caption         =   "Global Color Table Flag"
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblColorResolution 
      Caption         =   "Color Resolution"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblSortFlag 
      Caption         =   "Sort Flag"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblGlobalColorTableSize 
      Caption         =   "Global Color Table Size"
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblLogicalScreenHeight 
      Caption         =   "Logical Screen Height"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblLogicalScreenWidth 
      Caption         =   "Logical Screen Width"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblSignature 
      Caption         =   "Signature"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   2415
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
Attribute VB_Name = "frmGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type IMAGE_GIF_HEADER
    bSignature(0 To 2) As Byte
    bVersion(0 To 2) As Byte
    iLogicalScreenWidth As Integer
    iLogicalScreenHeight As Integer
    bPackedFields As Byte
    bBackgroundColorIndex As Byte
    bPixelAspectRatio As Byte
End Type

Const sLocation As String = "frmGIF"


Private Sub cmdGetInfo_Click()
On Error GoTo VB_Error

    chkSignature.value = 0
    txtVersion.Text = vbNullString
    txtLogicalScreenWidth.Text = vbNullString
    txtLogicalScreenHeight.Text = vbNullString
    chkGlobalColorTableFlag.value = 0
    txtColorResolution.Text = vbNullString
    chkSortFlag.value = 0
    txtGlobalColorTableSize.Text = vbNullString
    txtBackgroundColorIndex.Text = vbNullString
    txtPixelAspectRatio.Text = vbNullString
    
    
    If txtSelected.Text = vbNullString Then Exit Sub
    
    If File_Size_Name(txtSelected.Text) = 0 Then Exit Sub
    
    
    Dim sFileContents As String
    Dim IMAGE_GIF_HEADER As IMAGE_GIF_HEADER
    
    sFileContents = File_Read_Name(txtSelected.Text, Len(IMAGE_GIF_HEADER), 0)
    If Len(sFileContents) < Len(IMAGE_GIF_HEADER) Then Exit Sub
    
    
    Call MoveMemory(IMAGE_GIF_HEADER, ByVal sFileContents, Len(IMAGE_GIF_HEADER))
    
    
    If CharUpper(ByteArray_String(IMAGE_GIF_HEADER.bSignature())) <> "GIF" Then Exit Sub
    Dim sBinary As String
    
    With IMAGE_GIF_HEADER
        chkSignature.value = 1
        
        txtVersion.Text = ByteArray_String(.bVersion)
        txtLogicalScreenWidth.Text = int16_uint16(.iLogicalScreenWidth)
        txtLogicalScreenHeight.Text = int16_uint16(.iLogicalScreenHeight)
        
        sBinary = ltoa_(.bPackedFields, 2)
        chkGlobalColorTableFlag.value = Mid$(sBinary, 8, 1)
        txtColorResolution.Text = strtoul_(Mid$(sBinary, 5, 3), 2)
        chkSortFlag.value = Mid$(sBinary, 4, 1)
        txtGlobalColorTableSize.Text = strtoul_(Mid$(sBinary, 1, 3), 2)
        
        txtBackgroundColorIndex.Text = .bBackgroundColorIndex
        txtPixelAspectRatio.Text = .bPixelAspectRatio
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
