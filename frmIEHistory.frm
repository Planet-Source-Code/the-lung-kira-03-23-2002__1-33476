VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIEHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE History"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmIEHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwTypedUrls 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   3480
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   350
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblTypedURLs 
      Caption         =   "Type URLs"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmIEHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmIEHistory"


Private Sub cmdClearAll_Click()
On Error GoTo VB_Error

    Dim sValueName() As String
    Dim sData() As String
    Dim lDataType() As Long
    Dim lCount As Long
    
    lCount = Reg_EnumValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", sValueName(), sData(), lDataType())
    
    
    Dim lIncrement As Long
    For lIncrement = 0 To lCount - 1
        If lDataType(lIncrement) = REG_SZ Then
            Call Reg_DeleteValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", sValueName(lIncrement))
        End If
    Next lIncrement
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdClearAll_Click")
Resume Next
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo VB_Error
    
    Call ListView_Clear(lvwTypedUrls)

    Dim sValueName() As String
    Dim sData() As String
    Dim lDataType() As Long
    Dim lCount As Long
    
    lCount = Reg_EnumValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", sValueName(), sData(), lDataType())
    
    Dim lIncrement As Long
    For lIncrement = 0 To lCount - 1
        If lDataType(lIncrement) = REG_SZ Then
            lvwTypedUrls.ListItems.Add , , sData(lIncrement)
        End If
    Next lIncrement
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdRefresh_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    lvwTypedUrls.ColumnHeaders.Add , , "Typed Urls"
    
    Call cmdRefresh_Click
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub
