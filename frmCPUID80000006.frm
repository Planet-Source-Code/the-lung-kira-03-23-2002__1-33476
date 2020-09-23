VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCPUID80000006 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPUID 80000006"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmCPUID80000006.frx":0000
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
Attribute VB_Name = "frmCPUID80000006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmCPUID80000006"


Private Sub Form_Load()
On Error GoTo VB_Error

    With lvwCacheTLB.ColumnHeaders
        .Add , , "Description"
        .Add , , "Value"
    End With
    
    
    If CPUIDLevelExt_MAX > strtoul_("80000005", 16) Then
        lblEAX.Enabled = False
        lblEBX.Enabled = False
        lblECX.Enabled = False
        lblEDX.Enabled = False
        lblCacheTLB.Enabled = False
        lvwCacheTLB.Enabled = False
        Exit Sub
    End If
    
    
    'EAX = 80000006
    
    Dim sRegister As String
    
    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    
    Call cpuid_(strtoul_("80000006", 16), outEAX, outEBX, outECX, outEDX)
    
    
    
    With lvwCacheTLB.ListItems
        .Add , , "L1 Cache And TLB Configuration Descriptors"
        .Add(, , "Code TLB Entries").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 1, 12)), 2)
        
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEAX, 2), 32))
        
        If Right$(sRegister, 16) = "0000000000000000" Then
            .Add , , "(Unified) 4/2 MB L2 TLB Configuration Descriptor"
        Else
            .Add , , "4/2 MB L2 TLB Configuration Descriptor"
        End If
        .Add(, , "Code TLB Entries").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 1, 12)), 2)
        .Add(, , "Code TLB Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 13, 4)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 13, 4)), 2))
        .Add(, , "Data TLB Entries").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 17, 12)), 2)
        .Add(, , "Data TLB Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 29, 4)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 29, 4)), 2))
        .Add , , vbNullString
        
        
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outEBX, 2), 32))
        
        .Add , , "4 KB L2 TLB Configuration Descriptor"
        .Add(, , "Code TLB Entries").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 1, 12)), 2)
        .Add(, , "Code TLB Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 13, 4)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 13, 4)), 2))
        .Add(, , "Data TLB Entries").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 17, 12)), 2)
        .Add(, , "Data TLB Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 29, 4)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 29, 4)), 2))
        .Add , , vbNullString
        
        
        sRegister = StrReverse(Right$(String$(32, "0") & ltoa_(outECX, 2), 32))
        
        .Add , , "Unified L2 Cache Configuration Descriptor"
        .Add(, , "Unified L2 Cache Line Size In Bytes").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 1, 8)), 2)
        .Add(, , "Unified L2 Cache Lines Per Tag").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 9, 4)), 2)
        .Add(, , "Unified L2 Cache Associativity").SubItems(1) = StrReverse(Mid$(sRegister, 13, 4)) & " " & CacheTLB_Select(strtol_(StrReverse(Mid$(sRegister, 13, 4)), 2))
        .Add(, , "Unified L2 Cache Size In KBs").SubItems(1) = strtol_(StrReverse(Mid$(sRegister, 17, 4)), 2)
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


Private Function CacheTLB_Select(ByVal sValue As String) As String
On Error GoTo VB_Error

    Select Case sValue
        Case "0000": sValue = "L2 Off"
        Case "0001": sValue = "Direct Mapped"
        Case "0010": sValue = "2-Way"
        Case "0100": sValue = "4-Way"
        Case "0110": sValue = "8-Way"
        Case "1000": sValue = "16-Way"
        Case "1111": sValue = "Full"
        Case Else: sValue = "Unknown"
    End Select
    
    CacheTLB_Select = sValue
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\CacheTLB_Select")
Resume Next
End Function
