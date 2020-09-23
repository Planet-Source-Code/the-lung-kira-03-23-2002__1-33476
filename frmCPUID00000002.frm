VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCPUID00000002 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPUID 00000002"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmCPUID00000002.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQueriesRequired 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
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
      Left            =   4560
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
      Left            =   4560
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
      Left            =   4560
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin MSComctlLib.ListView lvwCacheTLB 
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2355
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
   Begin VB.Label lblQueriesRequired 
      Caption         =   "Queries Required"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblCacheTLB 
      Caption         =   "Cache - TLB"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
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
Attribute VB_Name = "frmCPUID00000002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmCPUID00000002"


Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwCacheTLB.ColumnHeaders
        .Add , , "Value"
        .Add , , "Description"
    End With
    
    
    If Not CPUIDLevel_MAX > &H1 Then
        lblEAX.Enabled = False
        lblEBX.Enabled = False
        lblECX.Enabled = False
        lblEDX.Enabled = False
        lblQueriesRequired.Enabled = False
        lblCacheTLB.Enabled = False
        lvwCacheTLB.Enabled = False
        Exit Sub
    End If
    
    
    'EAX = 2
    
    Dim sRegister As String
    Dim lIncrement As Long
    Dim lQuery As Long
    
    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    
    Call cpuid_(2, outEAX, outEBX, outECX, outEDX)
    
    sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEAX, 2), 32))
    
    
    lQuery = strtol_(StrReverse(Mid$(sRegister, 1, 8)), 2)
    
    For lIncrement = 1 To lQuery
        Call cpuid_(2, outEAX, outEBX, outECX, outEDX)
        DoEvents
    Next lIncrement
    
    
    txtQueriesRequired.Text = lQuery
    
    With lvwCacheTLB.ListItems
        .Add(, , StrReverse(Mid$(sRegister, 9, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 9, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 17, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 17, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 25, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 25, 8)), 2))
        
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEBX, 2), 32))
        
        .Add(, , StrReverse(Mid$(sRegister, 1, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 1, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 9, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 9, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 17, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 17, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 25, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 25, 8)), 2))
        
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outECX, 2), 32))
        
        .Add(, , StrReverse(Mid$(sRegister, 1, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 1, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 9, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 9, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 17, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 17, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 25, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 25, 8)), 2))
        
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEDX, 2), 32))
        
        .Add(, , StrReverse(Mid$(sRegister, 1, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 1, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 9, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 9, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 17, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 17, 8)), 2))
        .Add(, , StrReverse(Mid$(sRegister, 25, 8))).SubItems(1) = CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 25, 8)), 2))
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

Private Function CacheTLB_Select(ByVal lValue As Long) As String
On Error GoTo VB_Error

    Dim sDescriptor As String
    
    Select Case lValue
        Case &H0: sDescriptor = "Null Descriptor"
        Case &H1: sDescriptor = "code TLB, 4K pages, 4 ways, 32 entries"
        Case &H2: sDescriptor = "code TLB, 4M pages, fully, 2 entries"
        Case &H3: sDescriptor = "data TLB, 4K pages, 4 ways, 64 entries"
        Case &H4: sDescriptor = "data TLB, 4M pages, 4 ways, 8 entries"
        Case &H6: sDescriptor = "code L1 cache, 8KB, 4 ways, 32 byte lines"
        Case &H8: sDescriptor = "code L1 cache, 16KB, 4 ways, 32 byte lines"
        Case &HA: sDescriptor = "data L1 cache, 8KB, 2 ways, 32 byte lines"
        Case &HC: sDescriptor = "data L1 cache, 16KB, 4 ways, 32 byte lines"
        Case &H22: sDescriptor = "code & data L3 cache, 512KB, 4 ways (!), 64 byte lines, sectored"
        Case &H23: sDescriptor = "code & data L3 cache, 1024KB, 8 ways, 64 byte lines, sectored"
        Case &H25: sDescriptor = "code & data L3 cache, 2048KB, 8 ways, 64 byte lines, sectored"
        Case &H29: sDescriptor = "code & data L3 cache, 4096KB, 8 ways, 64 byte lines, sectored"
        Case &H40: sDescriptor = "no integrated L2 cache (P6 core) or L3 cache (P4 core)"
        Case &H41: sDescriptor = "code & data L2 cache, 128KB, 4 ways, 32 byte lines"
        Case &H42: sDescriptor = "code & data L2 cache, 256KB, 4 ways, 32 byte lines"
        Case &H43: sDescriptor = "code & data L2 cache, 512KB, 4 ways, 32 byte lines"
        Case &H44: sDescriptor = "code & data L2 cache, 1024KB, 4 ways, 32 byte lines"
        Case &H45: sDescriptor = "code & data L2 cache, 2048KB, 4 ways, 32 byte lines"
        Case &H50: sDescriptor = "code TLB, 4K/4M/2M pages, fully, 64 entries"
        Case &H51: sDescriptor = "code TLB, 4K/4M/2M pages, fully, 128 entries"
        Case &H52: sDescriptor = "code TLB, 4K/4M/2M pages, fully, 256 entries"
        Case &H5B: sDescriptor = "data TLB, 4K/4M pages, fully, 64 entries"
        Case &H5C: sDescriptor = "data TLB, 4K/4M pages, fully, 128 entries"
        Case &H5D: sDescriptor = "data TLB, 4K/4M pages, fully, 256 entries"
        Case &H66: sDescriptor = "data L1 cache, 8KB, 4 ways, 64 byte lines, sectored"
        Case &H67: sDescriptor = "data L1 cache, 16KB, 4 ways, 64 byte lines, sectored"
        Case &H68: sDescriptor = "data L1 cache, 32KB, 4 ways, 64 byte lines, sectored"
        Case &H70: sDescriptor = "trace L1 cache, 12 KµOPs, 4 ways"
        Case &H71: sDescriptor = "trace L1 cache, 16 KµOPs, 4 ways"
        Case &H72: sDescriptor = "trace L1 cache, 32 KµOPs, 4 ways"
        Case &H79: sDescriptor = "code & data L2 cache, 128KB, 8 ways, 64 byte lines, sectored"
        Case &H7A: sDescriptor = "code & data L2 cache, 256KB, 8 ways, 64 byte lines, sectored"
        Case &H7B: sDescriptor = "code & data L2 cache, 512KB, 8 ways, 64 byte lines, sectored"
        Case &H7C: sDescriptor = "code & data L2 cache, 1024KB, 8 ways, 64 byte lines, sectored"
        Case &H81: sDescriptor = "code & data L2 cache, 128KB, 8 ways, 32 byte lines"
        Case &H82: sDescriptor = "code & data L2 cache, 256KB, 8 ways, 32 byte lines"
        Case &H83: sDescriptor = "code & data L2 cache, 512KB, 8 ways, 32 byte lines"
        Case &H84: sDescriptor = "code & data L2 cache, 1024KB, 8 ways, 32 byte lines"
        Case &H85: sDescriptor = "code & data L2 cache, 2048KB, 8 ways, 32 byte lines"
        Case Else: sDescriptor = "Unknown " & lValue
    End Select
    
    CacheTLB_Select = sDescriptor
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\CacheTLB_Select")
Resume Next
End Function
