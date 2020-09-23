VERSION 5.00
Begin VB.Form frmFileAttributes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Attributes"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmFileAttributes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox drvFileAttributes 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4800
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
   Begin VB.DirListBox dirFileAttributes 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   840
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
      Width           =   5895
   End
   Begin VB.CheckBox chkTemporary 
      Height          =   255
      Left            =   5760
      TabIndex        =   30
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkSystem 
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkSparseFile 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkReparsePoint 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   24
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkReadOnly 
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkOffline 
      Height          =   255
      Left            =   5760
      TabIndex        =   20
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkNotContentIndexed 
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkNormal 
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkHidden 
      Height          =   255
      Left            =   5760
      TabIndex        =   14
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkEncrypted 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkDirectory 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkCompressed 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkArchive 
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5040
      TabIndex        =   31
      Top             =   4080
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
   Begin VB.Label lblTemporary 
      Caption         =   "Temporary"
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblSystem 
      Caption         =   "System"
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblSparseFile 
      Caption         =   "Sparse File"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblReparsePoint 
      Caption         =   "Reparse Point"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblReadOnly 
      Caption         =   "Read Only"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblOffline 
      Caption         =   "Offline"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblNotContentIndexed 
      Caption         =   "Not Content Indexed"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblNormal 
      Caption         =   "Normal"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblHidden 
      Caption         =   "Hidden"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblEncrypted 
      Caption         =   "Encrypted"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblDirectory 
      Caption         =   "Directory"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblCompressed 
      Caption         =   "Compressed"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblArchive 
      Caption         =   "Archive"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmFileAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim lFileAttributes As Long
Const sLocation As String = "frmFileAttributes"


Private Sub chkNormal_Click()
On Error GoTo VB_Error
    
    If chkNormal.value = 1 Then
        chkArchive.value = 0
        chkHidden.value = 0
        chkNotContentIndexed.value = 0
        chkOffline.value = 0
        chkReadOnly.value = 0
        chkSystem.value = 0
        chkTemporary.value = 0
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkNormal_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error

    If txtSelected.Text <> vbNullString Then
        lFileAttributes = 0
        
        If chkNormal.value = 1 Then
            lFileAttributes = FILE_ATTRIBUTE_NORMAL
        Else
            If chkArchive.value = 1 Then lFileAttributes = lFileAttributes Or FILE_ATTRIBUTE_ARCHIVE
            If chkHidden.value = 1 Then lFileAttributes = lFileAttributes Or FILE_ATTRIBUTE_HIDDEN
            If chkNotContentIndexed.value = 1 Then lFileAttributes = lFileAttributes Or FILE_ATTRIBUTE_NOT_CONTENT_INDEXED
            If chkOffline.value = 1 Then lFileAttributes = lFileAttributes Or FILE_ATTRIBUTE_OFFLINE
            If chkReadOnly.value = 1 Then lFileAttributes = lFileAttributes Or FILE_ATTRIBUTE_READONLY
            If chkSystem.value = 1 Then lFileAttributes = lFileAttributes Or FILE_ATTRIBUTE_SYSTEM
            If chkTemporary.value = 1 Then lFileAttributes = lFileAttributes Or FILE_ATTRIBUTE_TEMPORARY
        End If
        
        If SetFileAttributes(txtSelected.Text, lFileAttributes) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetFileAttributes")
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub dirFileAttributes_Change()
On Error GoTo VB_Error

    fileFileAttributes.Path = dirFileAttributes.Path
    txtSelected.Text = Str_BckSlhTerm_Fix(dirFileAttributes.Path)
    Call ProcessAttributes(txtSelected.Text)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\dirFileAttributes_Change")
Resume Next
End Sub

Private Sub dirFileAttributes_Click()
On Error GoTo VB_Error

    fileFileAttributes.Path = dirFileAttributes.Path
    txtSelected.Text = Str_BckSlhTerm_Fix(dirFileAttributes.Path)
    Call ProcessAttributes(txtSelected.Text)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\dirFileAttributes_Click")
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
    Call ProcessAttributes(txtSelected.Text)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\fileFileAttributes_Click")
Resume Next
End Sub

Private Sub ProcessAttributes(strFileName As String)
On Error GoTo VB_Error
    
    lFileAttributes = GetFileAttributes(strFileName)
    If lFileAttributes = -1 Then
        Call Error_API(Err.LastDllError, sLocation & "\ProcessAttributes", "GetFileAttributes")
        cmdApply.Enabled = False
        lFileAttributes = 0
    Else
        cmdApply.Enabled = True
    End If
    
    chkArchive.value = IIf(lFileAttributes And FILE_ATTRIBUTE_ARCHIVE, 1, 0)
    chkCompressed.value = IIf(lFileAttributes And FILE_ATTRIBUTE_COMPRESSED, 1, 0)
    chkDirectory.value = IIf(lFileAttributes And FILE_ATTRIBUTE_DIRECTORY, 1, 0)
    chkEncrypted.value = IIf(lFileAttributes And FILE_ATTRIBUTE_ENCRYPTED, 1, 0)
    chkHidden.value = IIf(lFileAttributes And FILE_ATTRIBUTE_HIDDEN, 1, 0)
    chkNormal.value = IIf(lFileAttributes And FILE_ATTRIBUTE_NORMAL, 1, 0)
    chkNotContentIndexed.value = IIf(lFileAttributes And FILE_ATTRIBUTE_NOT_CONTENT_INDEXED, 1, 0)
    chkOffline.value = IIf(lFileAttributes And FILE_ATTRIBUTE_OFFLINE, 1, 0)
    chkReadOnly.value = IIf(lFileAttributes And FILE_ATTRIBUTE_READONLY, 1, 0)
    chkReparsePoint.value = IIf(lFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT, 1, 0)
    chkSparseFile.value = IIf(lFileAttributes And FILE_ATTRIBUTE_SPARSE_FILE, 1, 0)
    chkSystem.value = IIf(lFileAttributes And FILE_ATTRIBUTE_SYSTEM, 1, 0)
    chkTemporary.value = IIf(lFileAttributes And FILE_ATTRIBUTE_TEMPORARY, 1, 0)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\ProcessAttributes")
Resume Next
End Sub
