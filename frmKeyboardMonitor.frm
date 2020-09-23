VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKeyboardMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Monitor"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmKeyboardMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Timer tmrKeyboardMonitor 
      Interval        =   2000
      Left            =   1680
      Top             =   1680
   End
   Begin MSComctlLib.ListView lvwKeyboardMonitor 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblMenu 
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   2055
   End
End
Attribute VB_Name = "frmKeyboardMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmKeyboardMonitor"


Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwKeyboardMonitor.ColumnHeaders
        .Add , , "Key Name"
        .Add , , "Count"
    End With
    
    With lvwKeyboardMonitor.ListItems
        .Add , , "Unknown"
        .Add , , "Left mouse button"
        .Add , , "Right mouse button"
        .Add , , "Control-break processing"
        .Add , , "Middle mouse button"
        .Add , , "X1 mouse button"
        .Add , , "X2 mouse button"
        .Add , , "Undefined"
        .Add , , "BACKSPACE key"
        .Add , , "TAB key"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "CLEAR key"
        .Add , , "ENTER key"
        .Add , , "Undefined"
        .Add , , "Undefined"
        .Add , , "SHIFT key"
        .Add , , "CTRL key"
        .Add , , "ALT key"
        .Add , , "PAUSE key"
        .Add , , "CAPS LOCK key"
        .Add , , "IME Kana/Hanguel/Hangul mode"
        .Add , , "Undefined"
        .Add , , "IME Junja mode"
        .Add , , "IME final mode"
        .Add , , "IME Hanja/Kanji mode"
        .Add , , "Undefined"
        .Add , , "ESC key"
        .Add , , "IME convert"
        .Add , , "IME nonconvert"
        .Add , , "IME accept"
        .Add , , "IME mode change request"
        .Add , , "SPACEBAR"
        .Add , , "PAGE UP key"
        .Add , , "PAGE DOWN key"
        .Add , , "END key"
        .Add , , "HOME key"
        .Add , , "LEFT ARROW key"
        .Add , , "UP ARROW key"
        .Add , , "RIGHT ARROW key"
        .Add , , "DOWN ARROW key"
        .Add , , "SELECT key"
        .Add , , "PRINT key"
        .Add , , "EXECUTE key"
        .Add , , "PRINT SCREEN key"
        .Add , , "INS key"
        .Add , , "DEL key"
        .Add , , "HELP key"
        .Add , , "0 key"
        .Add , , "1 key"
        .Add , , "2 key"
        .Add , , "3 key"
        .Add , , "4 key"
        .Add , , "5 key"
        .Add , , "6 key"
        .Add , , "7 key"
        .Add , , "8 key"
        .Add , , "9 key"
        .Add , , "Undefined"
        .Add , , "Undefined"
        .Add , , "Undefined"
        .Add , , "Undefined"
        .Add , , "Undefined"
        .Add , , "Undefined"
        .Add , , "Undefined"
        .Add , , "A key"
        .Add , , "B key"
        .Add , , "C key"
        .Add , , "D key"
        .Add , , "E key"
        .Add , , "F key"
        .Add , , "G key"
        .Add , , "H key"
        .Add , , "I key"
        .Add , , "J key"
        .Add , , "K key"
        .Add , , "L key"
        .Add , , "M key"
        .Add , , "N key"
        .Add , , "O key"
        .Add , , "P key"
        .Add , , "Q key"
        .Add , , "R key"
        .Add , , "S key"
        .Add , , "T key"
        .Add , , "U key"
        .Add , , "V key"
        .Add , , "W key"
        .Add , , "X key"
        .Add , , "Y key"
        .Add , , "Z key"
        .Add , , "Left Windows key"
        .Add , , "Right Windows key"
        .Add , , "Applications key"
        .Add , , "Reserved"
        .Add , , "Computer Sleep key"
        .Add , , "Numeric keypad 0 key"
        .Add , , "Numeric keypad 1 key"
        .Add , , "Numeric keypad 2 key"
        .Add , , "Numeric keypad 3 key"
        .Add , , "Numeric keypad 4 key"
        .Add , , "Numeric keypad 5 key"
        .Add , , "Numeric keypad 6 key"
        .Add , , "Numeric keypad 7 key"
        .Add , , "Numeric keypad 8 key"
        .Add , , "Numeric keypad 9 key"
        .Add , , "Multiply key"
        .Add , , "Add key"
        .Add , , "Separator key"
        .Add , , "Subtract key"
        .Add , , "Decimal key"
        .Add , , "Divide key"
        .Add , , "F1 key"
        .Add , , "F2 key"
        .Add , , "F3 key"
        .Add , , "F4 key"
        .Add , , "F5 key"
        .Add , , "F6 key"
        .Add , , "F7 key"
        .Add , , "F8 key"
        .Add , , "F9 key"
        .Add , , "F10 key"
        .Add , , "F11 key"
        .Add , , "F12 key"
        .Add , , "F13 key"
        .Add , , "F14 key"
        .Add , , "F15 key"
        .Add , , "F16 key"
        .Add , , "F17 key"
        .Add , , "F18 key"
        .Add , , "F19 key"
        .Add , , "F20 key"
        .Add , , "F21 key"
        .Add , , "F22 key"
        .Add , , "F23 key"
        .Add , , "F24 key"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "NUM LOCK key"
        .Add , , "SCROLL LOCK key"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Left SHIFT key"
        .Add , , "Right SHIFT key"
        .Add , , "Left CONTROL key"
        .Add , , "Right CONTROL key"
        .Add , , "Left MENU key"
        .Add , , "Right MENU key"
        .Add , , "Browser Back key"
        .Add , , "Browser Forward key"
        .Add , , "Browser Refresh key"
        .Add , , "Browser Stop key"
        .Add , , "Browser Search key"
        .Add , , "Browser Favorites key"
        .Add , , "Browser Start and Home key"
        .Add , , "Volume Mute key"
        .Add , , "Volume Down key"
        .Add , , "Volume Up key"
        .Add , , "Next Track key"
        .Add , , "Previous Track key"
        .Add , , "Stop Media key"
        .Add , , "Play/Pause Media key"
        .Add , , "Start Mail key"
        .Add , , "Select Media key"
        .Add , , "Start Application 1 key"
        .Add , , "Start Application 2 key"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "OEM 1"
        .Add , , "OEM PLUS"
        .Add , , "OEM COMMA"
        .Add , , "OEM MINUS"
        .Add , , "OEM PERIOD"
        .Add , , "OEM 2"
        .Add , , "OEM 3"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Reserved"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "Unassigned"
        .Add , , "OEM 4"
        .Add , , "OEM 5"
        .Add , , "OEM 6"
        .Add , , "OEM 7"
        .Add , , "OEM 8"
        .Add , , "Reserved"
        .Add , , "OEM specific"
        .Add , , "OEM 102"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "IME PROCESS key"
        .Add , , "OEM specific"
        .Add , , "PACKET"
        .Add , , "Unassigned"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "OEM specific"
        .Add , , "Attn key"
        .Add , , "CrSel key"
        .Add , , "ExSel key"
        .Add , , "Erase EOF key"
        .Add , , "Play key"
        .Add , , "Zoom key"
        .Add , , "Reserved for future use"
        .Add , , "PA1 key"
        .Add , , "Clear key"
        .Add , , "Unknown"
    End With
    
    
    Forms_Loaded.bKeyboarMonitor = True
    
    
    Dim lKeys As Long
    For lKeys = 0 To 255
        lvwKeyboardMonitor.ListItems(lKeys + 1).SubItems(1) = KeyboardMonitor(lKeys)
    Next lKeys
    
    txtTotal.Text = FormatNumber(KeyboardMonitor(256), 0, , , True)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error
    
    Forms_Loaded.bKeyboarMonitor = False
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
