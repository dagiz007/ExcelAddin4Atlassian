VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJiraDownloaIssusAttachments 
   Caption         =   "Download Jira attachments"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910.001
   OleObjectBlob   =   "frmJiraDownloaIssusAttachments.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmJiraDownloaIssusAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim jiraClient As New JiraRestClient
    Dim issues As Collection
    Dim issue As issue
    Dim attachment As attachment
    Dim counter As Integer
    
    Dim jiraAttachmentsDownloadFolder As String
    jiraAttachmentsDownloadFolder = GetSetting("ExcelAddIn4Jira", "Settings", "Jira_attachments_download_folder")
    
    Set issues = jiraClient.getJiraIssues(txtJQL)
    
    If Dir(jiraAttachmentsDownloadFolder, vbDirectory) = Empty Then MkDir jiraAttachmentsDownloadFolder
    
    For Each issue In issues
        
        counter = 1
        
        For Each attachment In issue.attachment
            Call jiraClient.saveJiraAttachmentToFile(attachment.id, jiraAttachmentsDownloadFolder & "\" & issue.jiraKey & "_" & counter & "_" & attachment.fileName)
            counter = counter + 1
        Next
        
    Next
        
    SaveSetting "ExcelAddIn4Jira", "DownloadJiraAttachments", "Jql", Trim(txtJQL)
    
    MsgBox "Vedleggene er nå lastet ned i " & jiraAttachmentsDownloadFolder
        
    Unload Me
    
End Sub

Private Sub lblExcelAddin4Jira_Click()
    ActiveWorkbook.FollowHyperlink "https://bitbucket.org/Stenstad/exceladd-in4jira/src/master/"
End Sub

Private Sub UserForm_Initialize()
    txtJQL = GetSetting("ExcelAddIn4Jira", "DownloadJiraAttachments", "Jql")
End Sub
