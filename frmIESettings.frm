VERSION 5.00
Begin VB.Form frmIESettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE Settings"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmIESettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkNotifyDownloadComplete 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkFriendlyHTTPErrors 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkGoButton 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtSearchPage 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox txtDefaultURL 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.TextBox txtDefaultSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox txtWindowTitle 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   17
      Top             =   3000
      Width           =   3855
   End
   Begin VB.CheckBox chkPersistantLinksFolder 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apply"
      Height          =   350
      Left            =   5040
      TabIndex        =   18
      Top             =   3360
      Width           =   975
   End
   Begin VB.CheckBox chkRatings 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblNotifyDownloadComplete 
      Caption         =   "Notify Download Complete"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblFriendlyHTTPErrors 
      Caption         =   "Friendly HTTP Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblGoButton 
      Caption         =   "Go Button"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblSearchPage 
      Caption         =   "Search Page"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblDefaultURL 
      Caption         =   "Default URL"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblDefaultSearch 
      Caption         =   "Default Search"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblWindowTitle 
      Caption         =   "Window Title"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblPersistantLinksFolder 
      Caption         =   "Persistant Links Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblRatings 
      Caption         =   "Ratings"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1935
   End
End
Attribute VB_Name = "frmIESettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sLocation As String = "frmIESettings"


Private Sub chkFriendlyHTTPErrors_Click()
On Error GoTo VB_Error

    If lblFriendlyHTTPErrors.Enabled = False Then lblFriendlyHTTPErrors.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkFriendlyHTTPErrors_Click")
Resume Next
End Sub

Private Sub chkGoButton_Click()
On Error GoTo VB_Error

    If lblGoButton.Enabled = False Then lblGoButton.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkGoButton_Click")
Resume Next
End Sub

Private Sub chkNotifyDownloadComplete_Click()
On Error GoTo VB_Error

    If lblNotifyDownloadComplete.Enabled = False Then lblNotifyDownloadComplete.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkNotifyDownloadComplete_Click")
Resume Next
End Sub

Private Sub chkPersistantLinksFolder_Click()
On Error GoTo VB_Error

    If lblPersistantLinksFolder.Enabled = False Then lblPersistantLinksFolder.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkPersistantLinksFolder_Click")
Resume Next
End Sub

Private Sub chkRatings_Click()
On Error GoTo VB_Error

    If lblRatings.Enabled = False Then lblRatings.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\chkRatings_Click")
Resume Next
End Sub

Private Sub cmdApply_Click()
On Error GoTo VB_Error

    If lblDefaultURL.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Page_URL", txtDefaultURL.Text, REG_SZ)
    If lblDefaultSearch.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Search_URL", txtDefaultSearch.Text, REG_SZ)
    
    If lblFriendlyHTTPErrors.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Friendly http errors", IIf(chkFriendlyHTTPErrors.value, "yes", "no"), REG_SZ)
    End If
    If lblGoButton.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "ShowGoButton", IIf(chkGoButton.value, "yes", "no"), REG_SZ)
    End If
    If lblNotifyDownloadComplete.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "NotifyDownloadComplete", IIf(chkNotifyDownloadComplete.value, "yes", "no"), REG_SZ)
    End If
    If lblPersistantLinksFolder.Enabled = True Then
        Call Reg_Write(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", "LinksFolderName", IIf(chkPersistantLinksFolder.value, "Links", vbNullString), REG_SZ)
    End If
    If lblRatings.Enabled = True Then
        If chkRatings.value = 1 Then
            Call Reg_DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\policies\Ratings", "Key")
        End If
    End If
    
    If lblSearchPage.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Search Page", txtSearchPage.Text, REG_SZ)
    If lblWindowTitle.Enabled = True Then Call Reg_Write(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Window Title", txtWindowTitle.Text, REG_SZ)
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\cmdApply_Click")
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo VB_Error

    Dim bFail As Byte
    Dim sReturn As String
    
    
    sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Page_URL", bFail)
    If bFail <> 0 Then
        lblDefaultURL.Enabled = False
    Else
        txtDefaultURL.Text = sReturn
    End If
    
    sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Search_URL", bFail)
    If bFail <> 0 Then
        lblDefaultSearch.Enabled = False
    Else
        txtDefaultSearch.Text = sReturn
    End If
    
    sReturn = Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Friendly http errors", bFail)
    If bFail <> 0 Then
        lblFriendlyHTTPErrors.Enabled = False
    Else
        If LCase(sReturn) = "yes" Then chkFriendlyHTTPErrors.value = 1
    End If
    
    sReturn = Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "ShowGoButton", bFail)
    If bFail <> 0 Then
        lblGoButton.Enabled = False
    Else
        If LCase(sReturn) = "yes" Then chkGoButton.value = 1
    End If
    
    sReturn = Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "NotifyDownloadComplete", bFail)
    If bFail <> 0 Then
        lblNotifyDownloadComplete.Enabled = False
    Else
        If LCase(sReturn) = "yes" Then chkNotifyDownloadComplete.value = 1
    End If
    
    sReturn = Reg_Read(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", "LinksFolderName", bFail)
    If bFail <> 0 Then
        lblPersistantLinksFolder.Enabled = False
    Else
        If sReturn <> vbNullString Then chkPersistantLinksFolder.value = 1
    End If
    
    sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\policies\Ratings", "Key", bFail)
    If bFail <> 0 Then
        lblRatings.Enabled = False
    Else
        If sReturn <> vbNullString Then chkRatings.value = 1
    End If
    
    sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Search Page", bFail)
    If bFail <> 0 Then
        lblSearchPage.Enabled = False
    Else
        txtSearchPage.Text = sReturn
    End If
    
    sReturn = Reg_Read(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Window Title", bFail)
    If bFail <> 0 Then
        lblWindowTitle.Enabled = False
    Else
        txtWindowTitle.Text = sReturn
    End If
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Form_Load")
Resume Next
End Sub

Private Sub txtDefaultSearch_Change()
On Error GoTo VB_Error

    If lblDefaultSearch.Enabled = False Then lblDefaultSearch.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtDefaultSearch_Change")
Resume Next
End Sub

Private Sub txtDefaultURL_Change()
On Error GoTo VB_Error

    If lblDefaultURL.Enabled = False Then lblDefaultURL.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtDefaultURL_Change")
Resume Next
End Sub

Private Sub txtSearchPage_Change()
On Error GoTo VB_Error

    If lblSearchPage.Enabled = False Then lblSearchPage.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtSearchPage_Change")
Resume Next
End Sub

Private Sub txtWindowTitle_Change()
On Error GoTo VB_Error

    If lblWindowTitle.Enabled = False Then lblWindowTitle.Enabled = True
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\txtWindowTitle_Change")
Resume Next
End Sub
