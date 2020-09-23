VERSION 5.00
Begin VB.Form frmKeyboardInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Info"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmKeyboardInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLayoutName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtLayoutID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtFunctionKeys 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtSubType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblLayoutID 
      Caption         =   "Layout ID"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblType 
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblSubType 
      Caption         =   "Sub Type"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblFunctionKeys 
      Caption         =   "Number of Function Keys"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLayoutName 
      Caption         =   "Layout Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmKeyboardInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmKeyboardInfo"


Private Sub Form_Load()
On Error GoTo VB_Error

    Select Case GetKeyboardType(2)
        Case 0
            Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "GetKeyboardType")
            txtFunctionKeys.Text = "Unknown"
        Case 1: txtFunctionKeys.Text = "10"
        Case 2: txtFunctionKeys.Text = "12/18"
        Case 3: txtFunctionKeys.Text = "10"
        Case 4: txtFunctionKeys.Text = "12"
        Case 5: txtFunctionKeys.Text = "10"
        Case 6: txtFunctionKeys.Text = "24"
        Case 7: txtFunctionKeys.Text = "10"
        Case Else: txtFunctionKeys.Text = "Hardware dependent and specified by the OEM"
    End Select
    
    txtLayoutName.Text = LangIdent(LOWORD(GetKeyboardLayout(0)))
    
    Dim sKeyboardLayout As String * KL_NAMELENGTH
    If GetKeyboardLayoutName(sKeyboardLayout) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "GetKeyboardLayoutName")
    txtLayoutID.Text = Right$(Str_NullTerm_Fix(sKeyboardLayout), 4)
    
    txtSubType.Text = GetKeyboardType(1)
    
    
    Dim lKeyboardType As Long
    lKeyboardType = GetKeyboardType(0)
    
    Select Case lKeyboardType
        Case 0: txtType.Text = "Unknown / Not Specified"
        Case 1: txtType.Text = "IBM PC/XT ( ) or compatible (83-key)"
        Case 2: txtType.Text = "Olivetti ICO (102-key) keyboard"
        Case 3: txtType.Text = "IBM PC/AT (84-key) or similar"
        Case 4: txtType.Text = "IBM enhanced (101- or 102-key)"
        Case 5: txtType.Text = "Nokia 1050 and similar"
        Case 6: txtType.Text = "Nokia 9140 and similar"
        Case 7: txtType.Text = "Japanese"
        Case Else: txtType.Text = "Unknown " & lKeyboardType
    End Select
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
