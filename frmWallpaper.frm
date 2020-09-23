VERSION 5.00
Begin VB.Form frmWallpaper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wallpaper"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmWallpaper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboDisplay 
      Height          =   315
      Left            =   3600
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ListBox lstNewWallpaper 
      Height          =   1035
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
   End
   Begin VB.TextBox txtCurrent 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   4320
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3360
      TabIndex        =   7
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image imgPreview 
      Height          =   1455
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblPreview 
      Caption         =   "Preview"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblDisplay 
      Caption         =   "Display"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblNewWallpaper 
      Caption         =   "New Wallpaper"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblCurrent 
      Caption         =   "Current"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmWallpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmWallpaper"


Private Sub cmdApply_Click()
On Error GoTo VB_Error

    If cboDisplay.ListIndex > -1 Then
        Call Reg_Write(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", cboDisplay.ListIndex, REG_SZ)
        Call Reg_Write(HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", IIf((cboDisplay.ListIndex = 1), 1, 0), REG_SZ)
    End If
    
    Select Case lstNewWallpaper.List(lstNewWallpaper.ListIndex)
        Case "Default"
            Call Reg_Write(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", vbNullString, REG_SZ)
            If SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, ByVal 0&, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
            
            txtCurrent.Text = vbNullString
        Case "None"
            Call Reg_Write(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", vbNullString, REG_SZ)
            If SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, ByVal vbNullString, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
            
            txtCurrent.Text = vbNullString
        Case Else
            Call Reg_Write(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", lstNewWallpaper.List(lstNewWallpaper.ListIndex), REG_SZ)
            If SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, ByVal lstNewWallpaper.List(lstNewWallpaper.ListIndex), SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
            
            txtCurrent.Text = lstNewWallpaper.List(lstNewWallpaper.ListIndex)
    End Select
    
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub cmdChoose_Click()
On Error GoTo VB_Error

    Dim sFileName As String

    Dim OPENFILENAME As OPENFILENAME
    With OPENFILENAME
        .flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_DONTADDTORECENT
        .hwndOwner = frmWallpaper.hwnd
        .lpstrFile = String$(MAX_PATH, 0)
        .lpstrFilter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar
        .lpstrTitle = "Open"
        .lStructSize = Len(OPENFILENAME)
        .nFilterIndex = 2
        .nMaxFile = Len(.lpstrFile)
    End With
    
    If GetOpenFileName(OPENFILENAME) = False Then
        Call Error_CommDlg(Err.LastDllError, sLocation & "\cmdChoose_Click", "GetOpenFileName")
    Else
        sFileName = Str_NullTerm_Fix(OPENFILENAME.lpstrFile)
        
        If File_Size_Name(sFileName) = 0 Then
            If MessageBoxEx(0&, "File size is 0.", "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdChoose_Click", "MessageBoxEx")
        Else
            lstNewWallpaper.AddItem sFileName
        End If
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdChoose_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With lstNewWallpaper
        .AddItem "Default"
        .AddItem "None"
    End With
    With cboDisplay
        .AddItem "Center"
        .AddItem "Tiled"
        .AddItem "Stretch"
    End With
    
    
    Dim sWallpaper As String
    sWallpaper = Reg_Read(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper")
    If sWallpaper <> vbNullString Then
        lstNewWallpaper.AddItem sWallpaper
        lstNewWallpaper.Selected(lstNewWallpaper.NewIndex) = True
        txtCurrent.Text = sWallpaper
    End If
    
    Dim lReturn As Long
    lReturn = Reg_Read(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle")
    Select Case lReturn
        Case 0 To 2: cboDisplay.ListIndex = lReturn
        Case Else: cboDisplay.ListIndex = 0
    End Select
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub lstNewWallpaper_Click()
On Error GoTo VB_Error

    imgPreview.Picture = Nothing
    
    If lstNewWallpaper.ListIndex > 1 Then
        If File_Exist(lstNewWallpaper.List(lstNewWallpaper.ListIndex)) = True Then
            imgPreview.Picture = LoadPicture(lstNewWallpaper.List(lstNewWallpaper.ListIndex))
        End If
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\lstNewWallpaper_Click")
Resume Next
End Sub
