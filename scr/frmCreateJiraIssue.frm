VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateJiraIssue 
   Caption         =   "Create Jira Issue"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9690.001
   OleObjectBlob   =   "frmCreateJiraIssue.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateJiraIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private attachments As New Collection

Private Sub cboProject_Change()

    cboIssueTypes.Clear
    
    Dim issueTypes As Collection
    Dim issueType As clsJiraIssuetype
    Dim issueTypeExist As Boolean
    
    Set issueTypes = Jira.GetIssueTypes(cboProject.value)

    For Each issueType In issueTypes
        If issueType.subtask = False Then
            If issueType.name = GetSetting("VbaAddin4Atlassian", "Settings", "lastCreatedIssueType") Then issueTypeExist = True
            cboIssueTypes.AddItem issueType.name
        End If
    Next
       
       
    If issueTypeExist Then
        cboIssueTypes.value = GetSetting("VbaAddin4Atlassian", "Settings", "lastCreatedIssueType")
    Else
        cboIssueTypes.ListIndex = 0
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()

    If Trim(txtSummary) = vbNullString Then
        lblError.Visible = True
        txtSummary.SetFocus
        Exit Sub
    End If
        

    Dim key As String
    key = Jira.CreateIssue(cboProject, cboIssueTypes, txtSummary, txtDescription)
            
    If Not key = "" Then
        
        SaveSetting "VbaAddin4Atlassian", "Settings", "lastCreatedProject", cboProject
        SaveSetting "VbaAddin4Atlassian", "Settings", "lastCreatedIssueType", cboIssueTypes
        
        Dim contr As Control
        
        For Each contr In Me.Controls
            If TypeName(contr) = "CheckBox" Then
                If contr.GroupName = "Attachment" And contr.value = True Then
                    Call Jira.AddAttachment(key, attachments(Int(contr.Tag)).filename, attachments(Int(contr.Tag)).data)
                End If
            End If
        Next
        
        If MsgBox("Issue " & key & " has been successfully created. Would you like to open the issue in your browser?", vbInformation + vbYesNo) = vbYes Then
            Call openHyperlink(gstrAtlassianURL & "/browse/" & key)
        End If
    End If
    
    Unload Me
End Sub

Private Sub lblLink_Click()
    Call openHyperlink("https://github.com/dagiz007/ExcelAddin4Atlassian")
End Sub

Private Sub UserForm_Initialize()
        
    Dim projects As Collection
    Dim project As clsJiraProject

    Set projects = Jira.GetProject()
    Dim projectExist As Boolean
    
    For Each project In projects
        cboProject.AddItem project.key
        If project.key = GetSetting("VbaAddin4Atlassian", "Settings", "lastCreatedProject") Then projectExist = True
        cboProject.List(cboProject.ListCount - 1, 1) = project.name
    Next
    
    cboProject.TextColumn = 2
    
    
    If Application.name = "Outlook" Then
        Call getSelectedEmail
    
        If attachments.Count > 0 Then
            Dim attachment As clsJiraIssueAttachment
            
            Dim iCheckBoxTop As Integer
            iCheckBoxTop = 0
       
            For Each attachment In attachments
                FrameAttachment.Controls.Add "Forms.CheckBox.1", "chkAttachment_" & attachment.id
                Controls("chkAttachment_" & attachment.id).GroupName = "Attachment"
                Controls("chkAttachment_" & attachment.id).Tag = attachment.id
                Controls("chkAttachment_" & attachment.id).Top = iCheckBoxTop
                Controls("chkAttachment_" & attachment.id).Left = 2
                Controls("chkAttachment_" & attachment.id).Height = 18
                Controls("chkAttachment_" & attachment.id).Width = 172
                Controls("chkAttachment_" & attachment.id).WordWrap = False
                Controls("chkAttachment_" & attachment.id).Caption = attachment.filename
                Controls("chkAttachment_" & attachment.id).value = True
                iCheckBoxTop = iCheckBoxTop + 20
            Next
        
            If attachments.Count > 5 Then
               FrameAttachment.Scrollbars = fmScrollBarsVertical
               FrameAttachment.ScrollHeight = iCheckBoxTop
            End If
        End If
    End If
    
            
    If projectExist Then
        cboProject.value = GetSetting("VbaAddin4Atlassian", "Settings", "lastCreatedProject")
    Else
        cboProject.ListIndex = 0
    End If
    
End Sub

Private Sub getSelectedEmail()

    Dim objApp As Outlook.Application
    Set objApp = Application
    
    Dim GetCurrentItem As MailItem
    
    
    If Not (objApp.ActiveExplorer.CurrentView = "Compact" Or objApp.ActiveExplorer.CurrentView = "Komprimer") Then Exit Sub
            
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            If objApp.ActiveExplorer.Selection.Count > 0 Then
                Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
            Else
                Exit Sub
            End If
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
    
    
    With GetCurrentItem
        txtSummary = .Subject
        txtDescription = .responseText
    
        If .attachments.Count > 0 Then

            Dim attachment As clsJiraIssueAttachment
            
            Dim i As Integer
           
            Dim path As String: path = Environ("temp") & "\"
            Dim filename As String
             
            For i = 1 To .attachments.Count
                Set attachment = New clsJiraIssueAttachment
                
                attachment.id = i
                
                filename = .attachments.item(i).filename
                attachment.filename = filename
                
                .attachments.item(i).SaveAsFile path & filename
                attachment.data = ReadFile(path & filename)
      
                Kill path & filename
 
                attachments.Add attachment
             Next
         End If
     End With
     
     Set objApp = Nothing
     Set GetCurrentItem = Nothing
 
End Sub

Private Sub UserForm_Terminate()
    Set attachments = Nothing
End Sub
