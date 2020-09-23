VERSION 5.00
Begin VB.Form frmSoundSentry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound Sentry"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmSoundSentry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMS3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "MS"
      Top             =   3600
      Width           =   255
   End
   Begin VB.TextBox txtMS2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "MS"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txtMS1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "MS"
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox txtWindowsEffectDLL 
      Height          =   285
      Left            =   2280
      TabIndex        =   26
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtWindowsEffectDuration 
      Height          =   285
      Left            =   2280
      TabIndex        =   23
      Text            =   "0"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox cboWindowsEffect 
      Height          =   315
      Left            =   2280
      TabIndex        =   21
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtGraphicEffectRGB 
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Text            =   "0"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtGraphicEffectDuration 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboGraphicEffect 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtTextEffectRGB 
      Height          =   285
      Left            =   2280
      TabIndex        =   19
      Text            =   "0"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtTextEffectDuration 
      Height          =   285
      Left            =   2280
      TabIndex        =   16
      Text            =   "0"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox chkSoundSentryOn 
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkIndicator 
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2880
      TabIndex        =   27
      Top             =   4320
      Width           =   975
   End
   Begin VB.CheckBox chkAvailable 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.ComboBox cboTextEffect 
      Height          =   315
      Left            =   2280
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblWindowsEffectDLL 
      Caption         =   "Windows Effect DLL"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblWindowsEffectDuration 
      Caption         =   "Windows Effect Duration"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lblWindowsEffect 
      Caption         =   "Windows Effect"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblGraphicEffectRGB 
      Caption         =   "Graphic Effect RGB"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblGraphicEffectDuration 
      Caption         =   "Graphic Effect Duration"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblGraphicEffect 
      Caption         =   "Graphic Effect"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblTextEffectRGB 
      Caption         =   "Text Effect RGB"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblTextEffectDuration 
      Caption         =   "Text Effect Duration"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblSoundSentryOn 
      Caption         =   "Sound Sentry On"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblIndicator 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblTextEffect 
      Caption         =   "Text Effect"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "frmSoundSentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmSoundSentry"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtGraphicEffectDuration.Text = MinMax(Val(txtGraphicEffectDuration.Text), 0, 2147483647)
    txtGraphicEffectRGB.Text = MinMax(Val(txtGraphicEffectRGB.Text), 0, 2147483647)
    txtTextEffectDuration.Text = MinMax(Val(txtTextEffectDuration.Text), 0, 2147483647)
    txtTextEffectRGB.Text = MinMax(Val(txtTextEffectRGB.Text), 0, 2147483647)
    If Len(txtWindowsEffectDLL.Text) > MAX_PATH Then txtWindowsEffectDLL.Text = Left$(txtWindowsEffectDLL.Text, MAX_PATH)
    txtWindowsEffectDuration.Text = MinMax(Val(txtWindowsEffectDuration.Text), 0, 2147483647)
    
    
    Dim SOUNDSENTRY As SOUNDSENTRY
    With SOUNDSENTRY
        .cbSize = Len(SOUNDSENTRY)
        
        .dwFlags = .dwFlags Or SSF_AVAILABLE
        If chkIndicator.value = 1 Then .dwFlags = .dwFlags Or SERKF_INDICATOR
        If chkSoundSentryOn.value = 1 Then .dwFlags = .dwFlags Or SSF_SOUNDSENTRYON
        
        If WinVersion(0, -1, True) = True Then
            Select Case cboGraphicEffect.ListIndex
                Case 0: .iFSGrafEffect = SSGF_DISPLAY
                Case 1: .iFSGrafEffect = SSGF_NONE
            End Select
            .iFSGrafEffectColor = txtGraphicEffectRGB.Text
            .iFSGrafEffectMSec = txtGraphicEffectDuration.Text
        
            Select Case cboTextEffect.ListIndex
                Case 0: .iFSTextEffect = SSTF_BORDER
                Case 1: .iFSTextEffect = SSTF_CHARS
                Case 2: .iFSTextEffect = SSTF_DISPLAY
                Case 3: .iFSTextEffect = SSTF_NONE
            End Select
            .iFSTextEffectColorBits = txtTextEffectRGB.Text
            .iFSTextEffectMSec = txtTextEffectDuration.Text
            
            .iWindowsEffectMSec = txtWindowsEffectDuration.Text
        End If
        
        Select Case cboWindowsEffect.ListIndex
            Case 0: .iWindowsEffect = SSWF_CUSTOM
            Case 1: .iWindowsEffect = SSWF_DISPLAY
            Case 2: .iWindowsEffect = SSWF_NONE
            Case 3: .iWindowsEffect = SSWF_TITLE
            Case 4: .iWindowsEffect = SSWF_WINDOW
        End Select
        .lpszWindowsEffectDLL = txtWindowsEffectDLL.Text
    End With
    
    If SystemParametersInfo(SPI_SETSOUNDSENTRY, SOUNDSENTRY.cbSize, SOUNDSENTRY, SPIF_UPDATEINIFILE) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SystemParametersInfo")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    With cboGraphicEffect
        .AddItem "Display"
        .AddItem "None"
    End With
    With cboTextEffect
        .AddItem "Flash Border"
        .AddItem "Flash Characters"
        .AddItem "Flash Display"
        .AddItem "None"
    End With
    With cboWindowsEffect
        .AddItem "Custom"
        .AddItem "Flash Display"
        .AddItem "None"
        .AddItem "Flash Title Bar"
        .AddItem "Flash Window"
    End With
    
    
    Dim SOUNDSENTRY As SOUNDSENTRY
    SOUNDSENTRY.cbSize = Len(SOUNDSENTRY)
    
    If SystemParametersInfo(SPI_GETSOUNDSENTRY, SOUNDSENTRY.cbSize, SOUNDSENTRY, 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "SystemParametersInfo")
    
    If SOUNDSENTRY.dwFlags And SSF_AVAILABLE Then
        With SOUNDSENTRY
            If .dwFlags And SSF_AVAILABLE Then chkAvailable.value = 1
            If .dwFlags And SSF_INDICATOR Then chkIndicator.value = 1
            If .dwFlags And SSF_SOUNDSENTRYON Then chkSoundSentryOn.value = 1
            
            
            Select Case .iWindowsEffect
                Case SSWF_CUSTOM: cboWindowsEffect.ListIndex = 0
                Case SSWF_DISPLAY: cboWindowsEffect.ListIndex = 1
                Case SSWF_NONE: cboWindowsEffect.ListIndex = 2
                Case SSWF_TITLE: cboWindowsEffect.ListIndex = 3
                Case SSWF_WINDOW: cboWindowsEffect.ListIndex = 4
                Case Else: cboWindowsEffect.ListIndex = -1
            End Select
            txtWindowsEffectDLL.Text = .lpszWindowsEffectDLL
            
            
            If WinVersion(0, -1, True) = True Then
                Select Case .iFSGrafEffect
                    Case SSGF_DISPLAY: cboGraphicEffect.ListIndex = 0
                    Case SSGF_NONE: cboGraphicEffect.ListIndex = 1
                    Case Else: cboGraphicEffect.ListIndex = -1
                End Select
                txtGraphicEffectDuration.Text = .iFSGrafEffectMSec
                txtGraphicEffectRGB.Text = .iFSGrafEffectColor
                
                Select Case .iFSTextEffect
                    Case SSTF_BORDER: cboTextEffect.ListIndex = 0
                    Case SSTF_CHARS: cboTextEffect.ListIndex = 1
                    Case SSTF_DISPLAY: cboTextEffect.ListIndex = 2
                    Case SSTF_NONE: cboTextEffect.ListIndex = 3
                    Case Else: cboTextEffect.ListIndex = -1
                End Select
                txtTextEffectDuration.Text = .iFSTextEffectMSec
                txtTextEffectRGB.Text = .iFSTextEffectColorBits
                
                txtWindowsEffectDuration.Text = .iWindowsEffectMSec
            Else
                lblGraphicEffect.Enabled = False
                cboGraphicEffect.Enabled = False
                lblGraphicEffectDuration.Enabled = False
                txtGraphicEffectDuration.Enabled = False
                lblGraphicEffectRGB.Enabled = False
                txtGraphicEffectRGB.Enabled = False
                lblTextEffect.Enabled = False
                cboTextEffect.Enabled = False
                lblTextEffectDuration.Enabled = False
                txtTextEffectDuration.Enabled = False
                lblTextEffectRGB.Enabled = False
                txtTextEffectRGB.Enabled = False
                lblWindowsEffectDuration.Enabled = False
                txtWindowsEffectDuration.Enabled = False
            End If
        End With
    Else
        lblIndicator.Enabled = False
        chkIndicator.Enabled = False
        lblSoundSentryOn.Enabled = False
        chkSoundSentryOn.Enabled = False
        lblGraphicEffect.Enabled = False
        cboGraphicEffect.Enabled = False
        lblGraphicEffectDuration.Enabled = False
        txtGraphicEffectDuration.Enabled = False
        lblGraphicEffectRGB.Enabled = False
        txtGraphicEffectRGB.Enabled = False
        lblTextEffect.Enabled = False
        cboTextEffect.Enabled = False
        lblTextEffectDuration.Enabled = False
        txtTextEffectDuration.Enabled = False
        lblTextEffectRGB.Enabled = False
        txtTextEffectRGB.Enabled = False
        lblWindowsEffect.Enabled = False
        cboWindowsEffect.Enabled = False
        lblWindowsEffectDLL.Enabled = False
        txtWindowsEffectDLL.Enabled = False
        lblWindowsEffectDuration.Enabled = False
        txtWindowsEffectDuration.Enabled = False
        
        cmdApply.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
