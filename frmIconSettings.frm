VERSION 5.00
Begin VB.Form frmIconSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Settings"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmIconSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFont 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "..."
      Height          =   225
      Left            =   3120
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtKB 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      TabIndex        =   11
      Text            =   "KB"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtIconCache 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Text            =   "0"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtShellIconDepth 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Text            =   "0"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtBPP 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Text            =   "BPP"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtPIX2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "PIX"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtPIX1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Text            =   "PIX"
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtHorizontalSpacing 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtVerticalSpacing 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2400
      TabIndex        =   17
      Top             =   2400
      Width           =   975
   End
   Begin VB.CheckBox chkTitleWrap 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblFont 
      Caption         =   "Font"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblIconCache 
      Caption         =   "Icon Cache"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblShellIconDepth 
      Caption         =   "Shell Icon Depth"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblVerticalSpacing 
      Caption         =   "Vertical Spacing"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblHorizontalSpacing 
      Caption         =   "Horizontal Spacing"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblTitleWrap 
      Caption         =   "Title Wrap"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmIconSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmIconSettings"

Dim LOGFONT As LOGFONT


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtHorizontalSpacing.Text = MinMax(Val(txtHorizontalSpacing.Text), 1, 65535)
    txtVerticalSpacing.Text = MinMax(Val(txtVerticalSpacing.Text), 1, 65535)
    txtShellIconDepth.Text = MinMax(Val(txtShellIconDepth.Text), 2, 1024)
    txtIconCache.Text = MinMax(Val(txtIconCache.Text), 0, 4294967295#)
    
    
    Dim ICONMETRICS As ICONMETRICS
    With ICONMETRICS
        .cbSize = Len(ICONMETRICS)
        .iHorzSpacing = txtHorizontalSpacing.Text
        .iVertSpacing = txtVerticalSpacing.Text
        .iTitleWrap = chkTitleWrap.value
        .lfFont = LOGFONT
    End With
    
    If SystemParametersInfo(SPI_SETICONMETRICS, Len(ICONMETRICS), ICONMETRICS, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
    
    If lblIconCache.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "Max Cached Icons", txtIconCache.Text, REG_SZ)
    If lblShellIconDepth.Enabled = True Then Call Reg_Write(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon BPP", txtShellIconDepth.Text, REG_SZ)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub cmdFont_Click()
On Error GoTo VB_Error

    Dim CHOOSEFONT_ As CHOOSEFONT_
    With CHOOSEFONT_
        .lStructSize = Len(CHOOSEFONT_)
        .hwndOwner = frmIconSettings.hwnd
        .hdc = frmIconSettings.hdc
        .lpLogFont = VarPtr(LOGFONT)
        .flags = CF_BOTH Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT
    End With

    If ChooseFont(CHOOSEFONT_) = False Then
        Call Error_CommDlg(Err.LastDllError, sLocation & "\cmdFont_Click", "ChooseFont")
    Else
        txtFont.Text = ByteArray_String(LOGFONT.lfFaceName())
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdFont_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim ICONMETRICS As ICONMETRICS
    ICONMETRICS.cbSize = Len(ICONMETRICS)
    
    If SystemParametersInfo(SPI_GETICONMETRICS, Len(ICONMETRICS), ICONMETRICS, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    
    With ICONMETRICS
        txtHorizontalSpacing.Text = .iHorzSpacing
        txtVerticalSpacing.Text = .iVertSpacing
        chkTitleWrap.value = IIf(.iTitleWrap, 1, 0)
        LOGFONT = .lfFont
        txtFont.Text = ByteArray_String(.lfFont.lfFaceName())
    End With
    
    
    Dim bFail As Byte
    
    txtIconCache.Text = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "Max Cached Icons", bFail)
    If bFail <> 0 Then lblIconCache.Enabled = False
    txtShellIconDepth.Text = Reg_Read(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon BPP", bFail)
    If bFail <> 0 Then lblShellIconDepth.Enabled = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub txtIconCache_Change()
On Error GoTo VB_Error

    If lblIconCache.Enabled = False Then lblIconCache.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtIconCache_Change")
Resume Next
End Sub

Private Sub txtShellIconDepth_Change()
On Error GoTo VB_Error

    If lblShellIconDepth.Enabled = False Then lblShellIconDepth.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtShellIconDepth_Change")
Resume Next
End Sub
