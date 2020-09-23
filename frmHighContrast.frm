VERSION 5.00
Begin VB.Form frmHighContrast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Contrast"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   Icon            =   "frmHighContrast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboDefaultScheme 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CheckBox chkIndicator 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeySound 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeyAvailable 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeyActive 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkHighContrastOn 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkConfirmHotKey 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkAvailable 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   1440
      TabIndex        =   16
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblDefaultScheme 
      Caption         =   "Default Scheme"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblIndicator 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblHotKeySound 
      Caption         =   "Hot Key Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblHotKeyAvailable 
      Caption         =   "Hot Key Available"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblHotKeyActive 
      Caption         =   "Hot Key Active"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblHighContrastOn 
      Caption         =   "High Contrast On"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblConfirmHotKey 
      Caption         =   "Confirm Hot Key"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmHighContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmHighContrast"


Private Sub cmdApply_Click()
On Error GoTo VB_Error

    Dim HIGHCONTRAST As HIGHCONTRAST
    With HIGHCONTRAST
        .cbSize = Len(HIGHCONTRAST)
        
        If chkConfirmHotKey.value = 1 Then .dwFlags = .dwFlags Or HCF_CONFIRMHOTKEY
        If chkHighContrastOn.value = 1 Then .dwFlags = .dwFlags Or HCF_HIGHCONTRASTON
        If chkHotKeyActive.value = 1 Then .dwFlags = .dwFlags Or HCF_HOTKEYACTIVE
        If chkHotKeyAvailable.value = 1 Then .dwFlags = .dwFlags Or HCF_HOTKEYAVAILABLE
        If chkHotKeySound.value = 1 Then .dwFlags = .dwFlags Or HCF_HOTKEYSOUND
        If chkIndicator.value = 1 Then .dwFlags = .dwFlags Or HCF_INDICATOR
        .dwFlags = .dwFlags Or HCF_AVAILABLE
        
        .lpszDefaultScheme = cboDefaultScheme.Text
    End With
    
    If SystemParametersInfo(SPI_SETHIGHCONTRAST, HIGHCONTRAST.cbSize, HIGHCONTRAST, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    Dim sValueName() As String
    Dim sData() As String
    Dim lDataType() As Long
    Dim lCount As Long
    
    lCount = Reg_EnumValue(HKEY_CURRENT_USER, "Control Panel\Appearance\Schemes", sValueName(), sData(), lDataType())
    
    Dim lIncrement As Long
    For lIncrement = 0 To lCount - 1
        cboDefaultScheme.AddItem sValueName(lIncrement)
    Next lIncrement
    
    
    Dim HIGHCONTRAST As HIGHCONTRAST
    HIGHCONTRAST.cbSize = Len(HIGHCONTRAST)
    
    If SystemParametersInfo(SPI_GETHIGHCONTRAST, HIGHCONTRAST.cbSize, HIGHCONTRAST, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    
    If HIGHCONTRAST.dwFlags And HCF_AVAILABLE Then
        With HIGHCONTRAST
            If (.dwFlags And HCF_AVAILABLE) Then chkAvailable.value = 1
            If (.dwFlags And HCF_CONFIRMHOTKEY) Then chkConfirmHotKey.value = 1
            If (.dwFlags And HCF_HIGHCONTRASTON) Then chkHighContrastOn.value = 1
            If (.dwFlags And HCF_HOTKEYACTIVE) Then chkHotKeyActive.value = 1
            If (.dwFlags And HCF_HOTKEYAVAILABLE) Then chkHotKeyAvailable.value = 1
            If (.dwFlags And HCF_HOTKEYSOUND) Then chkHotKeySound.value = 1
            If (.dwFlags And HCF_INDICATOR) Then chkIndicator.value = 1
            
            cboDefaultScheme.Text = .lpszDefaultScheme
        End With
    Else
        lblConfirmHotKey.Enabled = False
        chkConfirmHotKey.Enabled = False
        lblHighContrastOn.Enabled = False
        chkHighContrastOn.Enabled = False
        lblHotKeyActive.Enabled = False
        chkHotKeyActive.Enabled = False
        lblHotKeyAvailable.Enabled = False
        chkHotKeyAvailable.Enabled = False
        lblHotKeySound.Enabled = False
        chkHotKeySound.Enabled = False
        lblIndicator.Enabled = False
        chkIndicator.Enabled = False
        lblDefaultScheme.Enabled = False
        cboDefaultScheme.Enabled = False
        cmdApply.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
