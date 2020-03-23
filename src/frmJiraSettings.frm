VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJiraSettings 
   Caption         =   "Settings"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6570
   OleObjectBlob   =   "frmJiraSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmJiraSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
     
    lblJiraUrl.ForeColor = vbBlack
    lblJiraUsername.ForeColor = vbBlack
    lblJiraPassword.ForeColor = vbBlack
    
    If txtJiraUrl = vbNullString Then
        lblJiraUrl.ForeColor = vbRed
        txtJiraUrl.SetFocus
        Exit Sub
    End If
    
    If txtJiraUsername = vbNullString Then
        lblJiraUsername.ForeColor = vbRed
        txtJiraUsername.SetFocus
        Exit Sub
    End If

    If txtJiraPassword = vbNullString Then
        lblJiraPassword.ForeColor = vbRed
        txtJiraPassword.SetFocus
        Exit Sub
    End If
    
    SaveSetting "ExcelAddIn4Jira", "Settings", "Jira_url", txtJiraUrl
    SaveSetting "ExcelAddIn4Jira", "Settings", "Jira_username", txtJiraUsername
    SaveSetting "ExcelAddIn4Jira", "Settings", "Jira_password", txtJiraPassword
    SaveSetting "ExcelAddIn4Jira", "Settings", "Jira_remember_password", chkRemember
    
    SaveSetting "ExcelAddIn4Jira", "Settings", "Jira_attachments_download_folder", txtJiraAttachmentsDownloadFolder
    
    Unload Me

End Sub

Private Sub cmdSelectJiraAttachmentDownloadFolder_Click()
    Dim fldr As FileDialog

    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = -1 Then
            txtJiraAttachmentsDownloadFolder = .SelectedItems(1)
        End If
    End With

    Set fldr = Nothing
End Sub

Private Sub lblExcelAddin4Jira_Click()
    ActiveWorkbook.FollowHyperlink "https://github.com/DagAtleStenstad/ExcelAdd-in4Jira"
End Sub

Private Sub UserForm_Initialize()

    txtJiraUrl = GetSetting("ExcelAddIn4Jira", "Settings", "Jira_url")
    txtJiraUsername = GetSetting("ExcelAddIn4Jira", "Settings", "Jira_username")
    
    If GetSetting("ExcelAddIn4Jira", "Settings", "Jira_remember_password") = "True" Then
       chkRemember.value = "True"
       txtJiraPassword = GetSetting("ExcelAddin4Jira", "Settings", "Jira_password")
    End If
    
    If GetSetting("ExcelAddIn4Jira", "Settings", "Jira_attachments_download_folder") = vbNullString Then
        txtJiraAttachmentsDownloadFolder = "c:\JiraAttachments"
    Else
        txtJiraAttachmentsDownloadFolder = GetSetting("ExcelAddIn4Jira", "Settings", "Jira_attachments_download_folder")
    End If
    
End Sub
