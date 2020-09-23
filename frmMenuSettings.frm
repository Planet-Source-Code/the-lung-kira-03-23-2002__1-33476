VERSION 5.00
Begin VB.Form frmMenuSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Settings"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmMenuSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDropShadow 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtMS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "MS"
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkSelectionFade 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox chkFlatMenu 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkMenuAnimation 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.ComboBox cboDropAlignment 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox chkMenuFade 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtShowDelay 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Text            =   "0"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2040
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblDropShadow 
      Caption         =   "Drop Shadow"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblSelectionFade 
      Caption         =   "Selection Fade"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblFlatMenu 
      Caption         =   "Flat Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblMenuAnimation 
      Caption         =   "Menu Animation"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblMenuFade 
      Caption         =   "Menu Fade"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblDropAlignment 
      Caption         =   "Drop Alignment"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblShowDelay 
      Caption         =   "Show Delay"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frmMenuSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmMenuSettings"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtShowDelay.Text = MinMax(Val(txtShowDelay.Text), 0, 999)
    
    
    Dim bAlign As Boolean
    Select Case cboDropAlignment.ListIndex
        Case 0: bAlign = True
        Case 1: bAlign = False
    End Select
    If SystemParametersInfo(SPI_SETMENUDROPALIGNMENT, bAlign, 0&, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
    If WinVersion(-1, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETMENUFADE, 0&, ByVal CBool(chkMenuFade.value), SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        If SystemParametersInfo(SPI_SETSELECTIONFADE, 0&, ByVal CBool(chkSelectionFade.value), SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    If WinVersion(-1, 5010000, True) = True Then
        If SystemParametersInfo(SPI_SETDROPSHADOW, 0&, ByVal CBool(chkDropShadow.value), SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
        If SystemParametersInfo(SPI_SETFLATMENU, 0&, ByVal CBool(chkFlatMenu.value), SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    If WinVersion(4010000, 0, True) = True Then
        If SystemParametersInfo(SPI_SETMENUSHOWDELAY, ByVal CLng(txtShowDelay.Text), 0&, SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETMENUANIMATION, 0&, ByVal CBool(chkMenuAnimation.value), SPIF_UPDATEINIFILE) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With cboDropAlignment
        .AddItem "Left"
        .AddItem "Right"
    End With
    

    Dim bValue As Boolean
    Dim lValue As Long
    
    If SystemParametersInfo(SPI_GETMENUDROPALIGNMENT, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    cboDropAlignment.ListIndex = IIf(bValue, 0, 1)
    
    If WinVersion(-1, 5000000, True) = True Then
        If SystemParametersInfo(SPI_GETMENUFADE, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkMenuFade.value = IIf(bValue, 1, 0)
        
        If SystemParametersInfo(SPI_GETSELECTIONFADE, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkSelectionFade.value = IIf(bValue, 1, 0)
    Else
        lblMenuFade.Enabled = False
        chkMenuFade.Enabled = False
        
        lblSelectionFade.Enabled = False
        chkSelectionFade.Enabled = False
    End If
    If WinVersion(-1, 5010000, True) = True Then
        If SystemParametersInfo(SPI_GETDROPSHADOW, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkDropShadow.value = IIf(bValue, 1, 0)
        
        If SystemParametersInfo(SPI_GETFLATMENU, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkFlatMenu.value = IIf(bValue, 1, 0)
    Else
        lblDropShadow.Enabled = False
        chkDropShadow.Enabled = False
        
        lblFlatMenu.Enabled = False
        chkFlatMenu.Enabled = False
    End If
    If WinVersion(4010000, 0, True) = True Then
        If SystemParametersInfo(SPI_GETMENUSHOWDELAY, 0&, lValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        txtShowDelay.Text = lValue
    Else
        lblShowDelay.Enabled = False
        txtShowDelay.Enabled = False
    End If
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_GETMENUANIMATION, 0&, bValue, 0&) = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
        chkMenuAnimation.value = IIf(bValue, 1, 0)
    Else
        lblMenuAnimation.Enabled = False
        chkMenuAnimation.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
