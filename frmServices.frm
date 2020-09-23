VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Services"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmServices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwServices 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4260
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
   Begin VB.CommandButton cmdEnumerate 
      Caption         =   "Enumerate"
      Height          =   350
      Left            =   3600
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim bStopEnumeration As Boolean
Const sLocation As String = "frmServices"


Private Sub cmdEnumerate_Click()
On Error GoTo VB_Error

    cmdEnumerate.Enabled = False
    
    
    Dim lPorts As Long
    Dim lReturn As Long
    Dim sName As String
    
    Dim lName As Long
    Dim lProto As Long
    
    Dim servent As servent
    
    sName = String$(255, 0)
    
    For lPorts = 0 To 65535
        lReturn = getservbyport(lPorts, ByVal 0&)
        If lReturn > 0 Then
            If IsBadReadPtr(lReturn, Len(servent)) = True Then
                Call Error_API(Err.LastDllError, sLocation & "\cmdEnumerate_Click", "IsBadReadPtr")
            Else
                Call MoveMemory(servent, ByVal lReturn, Len(servent))
            End If
            If IsBadReadPtr(servent.s_name, 255) = True Then
                Call Error_API(Err.LastDllError, sLocation & "\cmdEnumerate_Click", "IsBadReadPtr")
            Else
                Call MoveMemory(ByVal sName, ByVal servent.s_name, 255)
            End If
            
            
            lName = InStr(1, sName, vbNullChar)
            lProto = InStr(lName + 1, sName, vbNullChar)
            
            With lvwServices.ListItems.Add(, , lPorts)
                .SubItems(1) = Mid$(sName, 1, lName - 1)
                .SubItems(2) = Mid$(sName, lName + 1, lProto)
            End With
            
            lblPercent.Caption = Percentage(lPorts, 65535, 0) & "%"
            
            DoEvents
        End If
        
        If bStopEnumeration = True Then Exit Sub
        If bShutdown = True Then Exit Sub
    Next lPorts
    
    lblPercent.Caption = vbNullString
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdEnumerate_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    If bWinsock = False Then
        lvwServices.Enabled = False
        cmdEnumerate.Enabled = False
    End If
    
    With lvwServices.ColumnHeaders
        .Add , , "Port"
        .Add , , "Service Name"
        .Add , , "Transport Protocol"
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo VB_Error

    bStopEnumeration = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Unload")
Resume Next
End Sub
