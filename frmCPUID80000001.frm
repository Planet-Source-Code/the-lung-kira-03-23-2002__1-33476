VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCPUID80000001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPUID 80000001"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmCPUID80000001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFamily 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtModel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtStepping 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtEDX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtECX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtEBX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtEAX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin MSComctlLib.ListView lvwFeatures 
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2143
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
   Begin VB.Label lblFeatures 
      Caption         =   "Features"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblFamily 
      Caption         =   "Family"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblModel 
      Caption         =   "Model"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblStepping 
      Caption         =   "Stepping"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblECX 
      Caption         =   "ECX"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblEDX 
      Caption         =   "EDX"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblEAX 
      Caption         =   "EAX"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblEBX 
      Caption         =   "EBX"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmCPUID80000001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmCPUID80000001"


Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwFeatures.ColumnHeaders
        .Add , , "Bit"
        .Add , , "Feature Name"
        .Add , , "Value"
    End With
    
    
    If Not CPUIDLevelExt_MAX > strtoul_("80000000", 16) Then
        lblEAX.Enabled = False
        lblEBX.Enabled = False
        lblECX.Enabled = False
        lblEDX.Enabled = False
        lblStepping.Enabled = False
        lblModel.Enabled = False
        lblFamily.Enabled = False
        lblFeatures.Enabled = False
        lvwFeatures.Enabled = False
        Exit Sub
    End If
    
    
    'EAX = 80000001
    
    Dim sRegister As String
    
    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    
    Call cpuid_(strtoul_("80000001", 16), outEAX, outEBX, outECX, outEDX)
    
    
    sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEAX, 2), 32))
    
    txtStepping.Text = strtol_(Right$("0000" & StrReverse(Mid$(sRegister, 1, 4)), 4), 2) & " - " & Right$("0000" & StrReverse(Mid$(sRegister, 1, 4)), 4)
    txtModel.Text = strtol_(Right$("0000" & StrReverse(Mid$(sRegister, 5, 4)), 4), 2) & " - " & Right$("0000" & StrReverse(Mid$(sRegister, 5, 4)), 4)
    txtFamily.Text = strtol_(Right$("0000" & StrReverse(Mid$(sRegister, 9, 4)), 4), 2) & " - " & Right$("0000" & StrReverse(Mid$(sRegister, 9, 4)), 4)
    
    
    sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEDX, 2), 32))
    
    With lvwFeatures.ListItems
        With .Add(, , "0")
            .SubItems(1) = "Floating Point Unit on chip"
            .SubItems(2) = CBool(Mid$(sRegister, 1, 1))
        End With
        With .Add(, , "1")
            .SubItems(1) = "Virtual 8086 Mode Extension"
            .SubItems(2) = CBool(Mid$(sRegister, 2, 1))
        End With
        With .Add(, , "2")
            .SubItems(1) = "Debugging Extension"
            .SubItems(2) = CBool(Mid$(sRegister, 3, 1))
        End With
        With .Add(, , "3")
            .SubItems(1) = "Page Size Extension"
            .SubItems(2) = CBool(Mid$(sRegister, 4, 1))
        End With
        With .Add(, , "4")
            .SubItems(1) = "Time Stamp Counter"
            .SubItems(2) = CBool(Mid$(sRegister, 5, 1))
        End With
        With .Add(, , "5")
            .SubItems(1) = "Model Specific Registers"
            .SubItems(2) = CBool(Mid$(sRegister, 6, 1))
        End With
        With .Add(, , "6")
            .SubItems(1) = "Physical Address Extension"
            .SubItems(2) = CBool(Mid$(sRegister, 7, 1))
        End With
        With .Add(, , "7")
            .SubItems(1) = "Machine Check Exception"
            .SubItems(2) = CBool(Mid$(sRegister, 8, 1))
        End With
        With .Add(, , "8")
            .SubItems(1) = "CMPXCHG8 Instruction"
            .SubItems(2) = CBool(Mid$(sRegister, 9, 1))
        End With
        With .Add(, , "9")
            .SubItems(1) = "On Chip APIC"
            .SubItems(2) = CBool(Mid$(sRegister, 10, 1))
        End With
        With .Add(, , "10")
            .SubItems(1) = "Reserved"
            .SubItems(2) = CBool(Mid$(sRegister, 11, 1))
        End With
        With .Add(, , "11")
            .SubItems(1) = "Fast System Call (SEP)"
            .SubItems(2) = CBool(Mid$(sRegister, 12, 1))
        End With
        With .Add(, , "12")
            .SubItems(1) = "Memory Type Range Registers"
            .SubItems(2) = CBool(Mid$(sRegister, 13, 1))
        End With
        With .Add(, , "13")
            .SubItems(1) = "Page Global Enable"
            .SubItems(2) = CBool(Mid$(sRegister, 14, 1))
        End With
        With .Add(, , "14")
            .SubItems(1) = "Machine Check Architecture"
            .SubItems(2) = CBool(Mid$(sRegister, 15, 1))
        End With
        With .Add(, , "15")
            .SubItems(1) = "Conditional Move and Compare Instructions"
            .SubItems(2) = CBool(Mid$(sRegister, 16, 1))
        End With
        With .Add(, , "16")
            .SubItems(1) = "Page Attribute Table"
            .SubItems(2) = CBool(Mid$(sRegister, 17, 1))
        End With
        With .Add(, , "17")
            .SubItems(1) = "36bit Page Size Extension"
            .SubItems(2) = CBool(Mid$(sRegister, 18, 1))
        End With
        With .Add(, , "18")
            .SubItems(1) = "Reserved"
            .SubItems(2) = CBool(Mid$(sRegister, 19, 1))
        End With
        With .Add(, , "19")
            .SubItems(1) = "Reserved"
            .SubItems(2) = CBool(Mid$(sRegister, 20, 1))
        End With
        With .Add(, , "20")
            .SubItems(1) = "Reserved"
            .SubItems(2) = CBool(Mid$(sRegister, 21, 1))
        End With
        With .Add(, , "21")
            .SubItems(1) = "Reserved"
            .SubItems(2) = CBool(Mid$(sRegister, 22, 1))
        End With
        With .Add(, , "22")
            .SubItems(1) = "MMX+ Technology"
            .SubItems(2) = CBool(Mid$(sRegister, 23, 1))
        End With
        With .Add(, , "23")
            .SubItems(1) = "MMX Technology"
            .SubItems(2) = CBool(Mid$(sRegister, 24, 1))
        End With
        With .Add(, , "24")
            .SubItems(1) = "Fast Save and Restor Instructions"
            .SubItems(2) = CBool(Mid$(sRegister, 25, 1))
        End With
        With .Add(, , "25")
            .SubItems(1) = "Reserved"
            .SubItems(2) = CBool(Mid$(sRegister, 26, 1))
        End With
        With .Add(, , "26")
            .SubItems(1) = "Reserved"
            .SubItems(2) = CBool(Mid$(sRegister, 27, 1))
        End With
        With .Add(, , "27")
            .SubItems(1) = "Reserved"
            .SubItems(2) = CBool(Mid$(sRegister, 28, 1))
        End With
        With .Add(, , "28")
            .SubItems(1) = "Reserved"
            .SubItems(2) = CBool(Mid$(sRegister, 29, 1))
        End With
        With .Add(, , "29")
            .SubItems(1) = "AA-64 Architecture"
            .SubItems(2) = CBool(Mid$(sRegister, 30, 1))
        End With
        With .Add(, , "30")
            .SubItems(1) = "3DNow!+ Technology"
            .SubItems(2) = CBool(Mid$(sRegister, 31, 1))
        End With
        With .Add(, , "31")
            .SubItems(1) = "3DNow! Technology"
            .SubItems(2) = CBool(Mid$(sRegister, 32, 1))
        End With
    End With
    
    
    txtEAX.Text = Right$("00000000" & ltoa_(outEAX, 16), 8)
    txtEBX.Text = Right$("00000000" & ltoa_(outEBX, 16), 8)
    txtECX.Text = Right$("00000000" & ltoa_(outECX, 16), 8)
    txtEDX.Text = Right$("00000000" & ltoa_(outEDX, 16), 8)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
