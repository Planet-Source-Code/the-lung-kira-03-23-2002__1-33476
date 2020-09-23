VERSION 5.00
Begin VB.Form frmFontSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Font Settings"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmFontSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboFontSmoothingType 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtFontSmoothingContrast 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2640
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.CheckBox chkFontSmoothing 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblFontSmoothingType 
      Caption         =   "Font Smoothing Type"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblFontSmoothingContrast 
      Caption         =   "Font Smoothing Contrast"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblFontSmoothing 
      Caption         =   "Font Smoothing"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmFontSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmFontSettings"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtFontSmoothingContrast.Text = MinMax(Val(txtFontSmoothingContrast.Text), 1000, 2200)
    
    
    If WinVersion(4010000, -1, False) = True Then
        Dim bValue As Boolean
        If SystemParametersInfo(SPI_GETWINDOWSEXTENSION, bValue, 0&, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        
        If bValue = True Then
            If SystemParametersInfo(SPI_SETFONTSMOOTHING, CBool(chkFontSmoothing.value), 0&, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        End If
    Else
        If SystemParametersInfo(SPI_SETFONTSMOOTHING, CBool(chkFontSmoothing.value), 0&, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    If WinVersion(-1, 5010000, True) = True Then
        If SystemParametersInfo(SPI_SETFONTSMOOTHINGCONTRAST, 0&, ByVal CLng(txtFontSmoothingContrast.Text), SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        
        Select Case cboFontSmoothingType.ListIndex
            Case 0: If SystemParametersInfo(SPI_SETFONTSMOOTHINGTYPE, 0&, ByVal FE_FONTSMOOTHINGSTANDARD, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
            Case 1: If SystemParametersInfo(SPI_SETFONTSMOOTHINGTYPE, 0&, ByVal FE_FONTSMOOTHINGCLEARTYPE, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        End Select
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With cboFontSmoothingType
        .AddItem "Standard"
        .AddItem "Clear Type"
    End With

    If WinVersion(4010000, -1, False) = True Then
        Dim byValue As Byte
        Dim bValue As Boolean
        byValue = 1
        
        If SystemParametersInfo(SPI_GETWINDOWSEXTENSION, byValue, 0&, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        If bValue = True Then
            If SystemParametersInfo(SPI_GETFONTSMOOTHING, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
            chkFontSmoothing.value = IIf(bValue, 1, 0)
        Else
            lblFontSmoothing.Enabled = False
            chkFontSmoothing.Enabled = False
        End If
    Else
        If SystemParametersInfo(SPI_GETFONTSMOOTHING, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkFontSmoothing.value = IIf(bValue, 1, 0)
    End If
    If WinVersion(-1, 5010000, True) = True Then
        Dim lValue As Long
        
        If SystemParametersInfo(SPI_GETFONTSMOOTHINGCONTRAST, 0&, lValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtFontSmoothingContrast.Text = lValue
        
        If SystemParametersInfo(SPI_GETFONTSMOOTHINGTYPE, 0&, lValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        Select Case lValue
            Case FE_FONTSMOOTHINGSTANDARD: cboFontSmoothingType.ListIndex = 0
            Case FE_FONTSMOOTHINGCLEARTYPE: cboFontSmoothingType.ListIndex = 1
            Case Else: cboFontSmoothingType.ListIndex = -1
        End Select
    Else
        lblFontSmoothingContrast.Enabled = False
        txtFontSmoothingContrast.Enabled = False
        
        lblFontSmoothingType.Enabled = False
        cboFontSmoothingType.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
