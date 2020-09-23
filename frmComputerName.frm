VERSION 5.00
Begin VB.Form frmComputerName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computer Name"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmComputerName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPhysicalDNSFullyQualified 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtPhysicalDNSDomain 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtPhysicalDNSHostname 
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtPhysicalNetBIOS 
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtDNSFullyQualified 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtDNSDomain 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtDNSHostname 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtNetBIOS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtComputerName 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   4200
      TabIndex        =   18
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblPhysicalDNSFullyQualified 
      Caption         =   "Physical DNS Fully Qualified"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblPhysicalDNSDomain 
      Caption         =   "Physical DNS Domain"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblPhysicalDNSHostname 
      Caption         =   "Physical DNS Hostname"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblPhysicalNetBIOS 
      Caption         =   "Physical Net BIOS"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblDNSFullyQualified 
      Caption         =   "DNS Fully Qualified"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblDNSDomain 
      Caption         =   "DNS Domain"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblDNSHostname 
      Caption         =   "DNS Hostname"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblNetBIOS 
      Caption         =   "Net BIOS"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblComputerName 
      Caption         =   "Computer Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmComputerName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmComputerName"


Private Sub cmdApply_Click()
On Error GoTo VB_Error
    
    txtComputerName.Text = Rem_NonStd_Chr(txtComputerName.Text)
    If Len(txtComputerName.Text) > MAX_COMPUTERNAME_LENGTH Then
        txtComputerName.Text = Left$(txtComputerName.Text, MAX_COMPUTERNAME_LENGTH)
    End If
    
    
    If SetComputerName(txtComputerName.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetComputerName")
    
    If WinVersion(-1, 5000000, True) = True Then
        txtPhysicalNetBIOS.Text = Rem_NonStd_Chr(txtPhysicalNetBIOS.Text)
        If Len(txtPhysicalNetBIOS.Text) > MAX_COMPUTERNAME_LENGTH Then
            txtPhysicalNetBIOS.Text = Left$(txtPhysicalNetBIOS.Text, MAX_COMPUTERNAME_LENGTH)
        End If
        txtPhysicalDNSHostname.Text = Rem_NonStd_Chr(txtPhysicalDNSHostname.Text)
        txtPhysicalDNSDomain.Text = Rem_NonStd_Chr(txtPhysicalDNSDomain.Text)
        
        If SetComputerNameEx(ComputerNamePhysicalNetBIOS, txtPhysicalNetBIOS.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetComputerNameEx")
        If SetComputerNameEx(ComputerNamePhysicalDnsHostname, txtPhysicalDNSHostname.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetComputerNameEx")
        If SetComputerNameEx(ComputerNamePhysicalDnsDomain, txtPhysicalDNSDomain.Text) = False Then Call Error_API(Err.LastDllError, sLocation & "\cmdApply_Click", "SetComputerNameEx")
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error
    
    txtComputerName.Text = ComputerName_Get
    
    If WinVersion(-1, 5000000, True) = True Then
        txtNetBIOS.Text = ComputerName_GetEx(ComputerNameNetBIOS)
        txtDNSHostname.Text = ComputerName_GetEx(ComputerNameDnsHostname)
        txtDNSDomain.Text = ComputerName_GetEx(ComputerNameDnsDomain)
        txtDNSFullyQualified.Text = ComputerName_GetEx(ComputerNameDnsFullyQualified)
        txtPhysicalNetBIOS.Text = ComputerName_GetEx(ComputerNamePhysicalNetBIOS)
        txtPhysicalDNSHostname.Text = ComputerName_GetEx(ComputerNamePhysicalDnsHostname)
        txtPhysicalDNSDomain.Text = ComputerName_GetEx(ComputerNamePhysicalDnsDomain)
        txtPhysicalDNSFullyQualified.Text = ComputerName_GetEx(ComputerNamePhysicalDnsFullyQualified)
    Else
        lblNetBIOS.Enabled = False
        txtNetBIOS.Enabled = False
        lblDNSHostname.Enabled = False
        lblDNSDomain.Enabled = False
        lblDNSFullyQualified.Enabled = False
        lblPhysicalNetBIOS.Enabled = False
        txtPhysicalNetBIOS.Enabled = False
        lblPhysicalDNSHostname.Enabled = False
        txtPhysicalDNSHostname.Enabled = False
        lblPhysicalDNSDomain.Enabled = False
        txtPhysicalDNSDomain.Enabled = False
        lblPhysicalDNSFullyQualified.Enabled = False
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub


Private Function ComputerName_GetEx(ByVal COMPUTER_NAME_FORMAT As COMPUTER_NAME_FORMAT) As String
On Error GoTo VB_Error
    
    Dim sComputerName As String
    sComputerName = String$(MAX_COMPUTERNAME_LENGTH * 2, 0)
    
    If GetComputerNameEx(COMPUTER_NAME_FORMAT, sComputerName, Len(sComputerName)) = False Then Call Error_API(Err.LastDllError, sLocation & "\Form_Load", "GetComputerNameEx")
    ComputerName_GetEx = sComputerName
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Function
        
Private Function Rem_NonStd_Chr(ByVal sData As String) As String
On Error GoTo VB_Error

    If sData = vbNullString Then Exit Function
    
    
    Dim lIncrement As Long
    
    For lIncrement = 0 To 32
        sData = Replace$(sData, Chr$(lIncrement), vbNullString, 1, -1)
    Next lIncrement
    For lIncrement = 42 To 44
        sData = Replace$(sData, Chr$(lIncrement), vbNullString, 1, -1)
    Next lIncrement
    For lIncrement = 58 To 63
        sData = Replace$(sData, Chr$(lIncrement), vbNullString, 1, -1)
    Next lIncrement
    For lIncrement = 91 To 93
        sData = Replace$(sData, Chr$(lIncrement), vbNullString, 1, -1)
    Next lIncrement
    For lIncrement = 128 To 255
        sData = Replace$(sData, Chr$(lIncrement), vbNullString, 1, -1)
    Next lIncrement
    
    
    sData = Replace$(sData, Chr$(34), vbNullString, 1, -1)
    sData = Replace$(sData, Chr$(47), vbNullString, 1, -1)
    sData = Replace$(sData, Chr$(96), vbNullString, 1, -1)
    sData = Replace$(sData, Chr$(124), vbNullString, 1, -1)
    
    Rem_NonStd_Chr = sData
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Rem_NonStd_Chr")
Resume Next
End Function
