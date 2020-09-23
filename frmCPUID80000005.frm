VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCPUID80000005 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPUID 80000005"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmCPUID80000005.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEDX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4200
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
      Left            =   4200
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
      Left            =   4200
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin MSComctlLib.ListView lvwCacheTLB 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   6615
      _ExtentX        =   11668
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
   Begin VB.Label lblCacheTLB 
      Caption         =   "Cache - TLB"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
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
   Begin VB.Label lblEAX 
      Caption         =   "EAX"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
   Begin VB.Label lblECX 
      Caption         =   "ECX"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmCPUID80000005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmCPUID80000005"


Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwCacheTLB.ColumnHeaders
        .Add , , "Description"
        .Add , , "Value"
    End With
    
    
    If Not CPUIDLevelExt_MAX > strtoul_("80000004", 16) Then
        lblEAX.Enabled = False
        lblEBX.Enabled = False
        lblECX.Enabled = False
        lblEDX.Enabled = False
        lblCacheTLB.Enabled = False
        lvwCacheTLB.Enabled = False
        Exit Sub
    End If
    
    
    'EAX = 80000005
    
    Dim sRegister As String
    
    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    
    Call cpuid_(strtoul_("80000005", 16), outEAX, outEBX, outECX, outEDX)
    
    
    With lvwCacheTLB.ListItems
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEAX, 2), 32))
        
        .Add , , "L1 Cache And TLB Configuration Descriptors"
        .Add(, , "Code TLB Entries").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 1, 12)), 2)
        .Add(, , "Code TLB Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 13, 4))
        .Add(, , "Data TLB Entries").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 17, 12)), 2)
        .Add(, , "Data TLB Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 29, 4))
        .Add , , vbNullString
        
        
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEBX, 2), 32))
        
        .Add , , "4 KB L1 TLB Configuration Descriptor"
        .Add(, , "Code TLB Entries").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 1, 12)), 2)
        .Add(, , "Code TLB Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 13, 4))
        .Add(, , "Data TLB Entries").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 17, 12)), 2)
        .Add(, , "Data TLB Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 29, 4))
        .Add , , vbNullString
        
        
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outECX, 2), 32))
        
        .Add , , "Data L1 Cache Configuration Descriptor"
        .Add(, , "Data L1 Cache Line Size In Bytes").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 1, 8)), 2)
        .Add(, , "Data L1 Cache Lines Per Tag").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 9, 4)), 2)
        .Add(, , "Data L1 Cache Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 13, 4))
        .Add(, , "Data L1 Cache Size In KBs").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 17, 16)), 2)
        .Add , , vbNullString
        
        
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEDX, 2), 32))
        
        .Add , , "Code L1 Cache Configuration Descriptor"
        .Add(, , "Code L1 Cache Line Size In Bytes").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 1, 8)), 2)
        .Add(, , "Code L1 Cache Lines Per Tag").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 9, 4)), 2)
        .Add(, , "Code L1 Cache Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 13, 4))
        .Add(, , "Code L1 Cache Size In KBs").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 17, 16)), 2)
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
