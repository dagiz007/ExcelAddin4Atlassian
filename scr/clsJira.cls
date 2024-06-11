VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private RestClient As New clsRestClient

Private Sub Class_initialize()
    If Not gblnSuccessfulLogin Then
        If Not CorrectCredentianls Then
            frmSettings.Show
        Else
            gblnSuccessfulLogin = True
        End If
    End If
End Sub

Private Sub Class_Terminate()
    Set RestClient = Nothing
End Sub

Public Function CorrectCredentianls() As Boolean
    If gstrAtlassianURL = vbNullString Then Call LoadSettings
    
    If gstrAtlassianURL = vbNullString Then
        CorrectCredentianls = False
    Else
        Set RestClient = RestClient.Service("GET", gstrAtlassianURL & "/rest/api/3/myself")
        If RestClient.status = 200 Then CorrectCredentianls = True
    End If
End Function

Public Function GetIssues(jql As String) As Collection

    Dim payload As String

    Dim json As Object
    Dim item As Variant
    Dim issue As clsJiraIssue
    Dim issues As New Collection
    Dim maxResults As Integer
    Dim startAt As Long
    Dim totalIssues As Long
    
    startAt = 0
    maxResults = 50
          
    Do
        DoEvents
                
        payload = "{""jql"": """ & Trim(jql) & """, ""startAt"": " & startAt & ",""maxResults"": " & maxResults & _
               ",""fields"": [""*all""], ""expand"": [""changelog,renderedFields""]}"
           
        'Version 2 use Jira Markup, but version 3 only support ADF (for Description and comment)
        Set RestClient = RestClient.Service("POST", gstrAtlassianURL & "/rest/api/2/search", payload)
        
        Set json = JsonConverter.ParseJson(RestClient.responseText)
        
        totalIssues = json("total")
        
        For Each item In json("issues")
            Set issue = New clsJiraIssue
            Set issue = GetIssue(item("key"), item)
            issues.Add issue
        Next
            
        startAt = startAt + maxResults
        
    Loop Until startAt >= totalIssues
    
    Set GetIssues = issues
    
End Function

Public Function CreateProject(key As String, name As String, leadAccountId As String, projectTemplateKey As String) As clsRestClient
    
    Dim payload As String

    payload = "{ ""assigneeType"": ""UNASSIGNED"", " & _
              " ""key"":""" & key & """," & _
              " ""name"":""" & name & """, " & _
              " ""leadAccountId"": """ & leadAccountId & """, " & _
              " ""projectTemplateKey"": """ & projectTemplateKey & """}"
             
    Set CreateProject = RestClient.Service("POST", gstrAtlassianURL & "/rest/api/3/project", payload)

End Function

Public Function AddActorToProjectRole(key As String, roleId As String, userType As String, Id As String) As clsRestClient
    
    'userType can be user or groupId
    
    Dim url As String
    url = gstrAtlassianURL & "/rest/api/3/project/" & key & "/role/" & roleId
    
    Dim payload As String
    payload = "{""" & userType & """: [ """ & Id & """  ]}"
             
    Set AddActorToProjectRole = RestClient.Service("POST", url, payload)

End Function

Public Function GetIssue(ByVal key As String, Optional ByVal json As Object) As clsJiraIssue

    On Error GoTo Errorhandler

    If gobjJiraIssueCache Is Nothing Then Set gobjJiraIssueCache = CreateObject("Scripting.Dictionary")
       
    'Check if issue is already cached
    If Not gobjJiraIssueCache.Exists(key) Then
        
        If json Is Nothing Then
            'Version 2 use Jira Markup, but version 3 only support ADF (for Description and comment)
            Set RestClient = RestClient.Service("GET", gstrAtlassianURL & "/rest/api/2/issue/" & key & "?expand=changelog,renderedFields")
            Set json = JsonConverter.ParseJson(RestClient.responseText)
        End If
        
        Dim issue As New clsJiraIssue
        
        issue.Id = json("id")
        issue.key = json("key")
        issue.summary = json("fields")("summary")
        issue.description = IIf(IsNull(json("fields")("description")), "", json("fields")("description"))
        issue.descriptionRendered = json("renderedFields")("description")
                
        If Not IsNull(json("fields")("assignee")) Then issue.assignee = AddUser(json("fields")("assignee"))
        If Not IsNull(json("fields")("reporter")) Then issue.reporter = AddUser(json("fields")("reporter"))
        If Not IsEmpty(json("fields")("creator")) Then issue.creator = AddUser(json("fields")("creator"))
     
        issue.issueType = AddIssueType(json("fields")("issuetype"))
        issue.project = AddIssueProject(json("fields")("project"))
        
        'We use changlog not transition, since transition dosent contain dates.
        issue.transition = AddIssueTransitions(json("changelog")("histories"))
        
        issue.comment = AddIssueComments(json("fields")("comment")("comments"))
        issue.attachment = AddAttachments(json("fields")("attachment"))
        issue.version = AddIssueFixVersions(json("fields")("fixVersions"))
        issue.createdDate = ParseISO8601date(json("fields")("created"))
        
        If Not IsNull(json("fields")("resolutiondate")) Then
            issue.resolutionDate = ParseISO8601date(json("fields")("resolutiondate"))
        End If
        
        issue.json = json
           
        gobjJiraIssueCache.Add key, issue
           
    End If
        
    Set GetIssue = gobjJiraIssueCache(key)
    
    Exit Function
    
Errorhandler:

    Debug.Print Err.Number & " " & Err.description
    Stop
    Resume Next
           
End Function

Public Function GetFormId(key As String) As Collection
    
    Dim json As Object
    Dim item As Variant
    
    Dim form As clsJiraForm
    Dim forms As New Collection
    
    Set RestClient = RestClient.Service("GET", "https://api.atlassian.com/jira/forms/cloud/" & getCloudId & "/issue/" & key & "/form", , True)
    
    Set json = JsonConverter.ParseJson(RestClient.responseText)
    
    For Each item In json
        Set form = New clsJiraForm
        form.Id = item("id")
        form.internal = item("internal")
        form.submitted = item("submitted")
        form.locked = item("lock")
        form.name = item("name")
        form.updated = ParseISO8601date(item("updated"))
            
        forms.Add form
    Next
    
    Set GetFormId = forms

End Function

Public Function GetForm(key As String, formId As String) As clsRestClient
    Set GetForm = RestClient.Service("GET", "https://api.atlassian.com/jira/forms/cloud/" & getCloudId & "/issue/" & key & "/form/" & formId, , True)
    Debug.Print GetForm.responseText
    'TODO write function
End Function

Public Function GetProjectRoles(projectId As String) As clsRestClient
    Set GetProjectRoles = RestClient.Service("GET", gstrAtlassianURL & "/rest/api/3/project/" & projectId & "/role/")
    Debug.Print GetProjectRoles.responseText
    'TODO write function
End Function

Public Function ReopenForm(key As String, formId As String) As clsRestClient
    Set ReopenForm = RestClient.Service("PUT", "https://api.atlassian.com/jira/forms/cloud/" & getCloudId & "/issue/" & key & "/form/" & formId & "/action/reopen", , True)
End Function

Public Function EditIssue(ByVal key As String, payload As String) As clsRestClient
    'Payload structure depending on the used api version. Version 2 use Jira Markup, but version 3 only support ADF (for updating description and comments).
    Set EditIssue = RestClient.Service("PUT", gstrAtlassianURL & "/rest/api/2/issue/" & key, payload)
End Function

Public Function UpdateComment(ByVal key As String, Id As Long, comment As String, Optional notifyUsers As Boolean = True) As clsRestClient
    
    Dim payload As String
    payload = "{""body"": " & ConvertToJson(comment) & ", ""notifyUsers"": """ & notifyUsers & """}"
    
    'Payload structure depending on the used api version. Version 2 use Jira Markup, but version 3 only support ADF.
    Set UpdateComment = RestClient.Service("PUT", gstrAtlassianURL & "/rest/api/2/issue/" & key & "/comment/" & Id, payload)
        
End Function

Public Function GetProjectRoleActor(projectId As String, roleId As String) As Collection
        
    DoEvents
    
    Dim json As Object
    Dim item As Variant
    
    Dim roleActor As clsJiraProjectRoleActor
    Dim roleActors As New Collection
    
    Set RestClient = RestClient.Service("GET", gstrAtlassianURL & "/rest/api/3/project/" & projectId & "/role/" & roleId)
    
    Set json = JsonConverter.ParseJson(RestClient.responseText)
    
    For Each item In json("actors")
        Set roleActor = New clsJiraProjectRoleActor
        
        'atlassian-user-role-actor, atlassian-group-role-actor
        roleActor.roleType = item("type")
        roleActor.name = item("displayName")
        
        If Not IsEmpty(item("actorUser")) Then
            roleActor.Id = item("actorUser")("accountId")
        Else
            roleActor.Id = item("actorGroup")("groupId")
        End If
        
        roleActors.Add roleActor
    Next
    
    Set GetProjectRoleActor = roleActors
    
End Function

Public Function GetProject() As Collection
    
    Dim json As Object
    Dim item As Variant
    
    Dim project As clsJiraProject
    Dim projects As New Collection
    
    Set RestClient = RestClient.Service("GET", gstrAtlassianURL & "/rest/api/3/project")
    
    Set json = JsonConverter.ParseJson(RestClient.responseText)
    
    For Each item In json
        Set project = New clsJiraProject
        
        project.Id = item("id")
        project.key = item("key")
        project.name = item("name")
        project.simplified = item("simplified")
                
        If Not IsEmpty(item("archived")) And item("archived") = True Then
            project.archived = True
        End If
        
        projects.Add project
    Next
    
    Set GetProject = projects

End Function

Public Function GetUser(accountId As String) As clsAtlassianUser
    
    Dim json As Object
   
    Set RestClient = RestClient.Service("GET", gstrAtlassianURL & "/rest/api/3/user?accountId=" & accountId)
    
    Set json = JsonConverter.ParseJson(RestClient.responseText)
    
    Set GetUser = AddUser(json)

End Function

Public Function GetUserFromGroup(groupId As String) As Collection
    
    Dim json As Object
    Dim item As Variant

    Set RestClient = RestClient.Service("GET", gstrAtlassianURL & "/rest/api/3/group/member?groupId=" & groupId)
    
    Dim user As clsAtlassianUser
    Dim users As New Collection

    Set json = JsonConverter.ParseJson(RestClient.responseText)
    
     For Each item In json("values")
        Set user = AddUser(item)
        users.Add user
    Next
    
    Set GetUserFromGroup = users

End Function

Public Function GetIssueTypes(jiraProjectId As String) As Collection
    
    Dim json As Object
    Dim item As Variant
    
    Dim issueType As clsJiraIssuetype
    Dim issueTypes As New Collection
        
    Set RestClient = RestClient.Service("GET", gstrAtlassianURL & "/rest/api/3/project/" & jiraProjectId & "/statuses")
    
    Set json = JsonConverter.ParseJson(RestClient.responseText)
    
     For Each item In json
        Set issueType = New clsJiraIssuetype
            
        issueType.Id = item("id")
        issueType.name = item("name")
        issueType.subtask = item("subtask")
        
        issueTypes.Add issueType
    Next
    
    Set GetIssueTypes = issueTypes
    
End Function

Public Function CreateIssue(project As String, issueType As String, summary As String, description As String) As String

    Dim payload As String

    payload = "{""fields"" : { " & _
        """project"": { ""key"": """ & project & """ }, " & _
        """summary"": " & ConvertToJson(summary) & ", "
        
        If Not Trim(description) = vbNullString Then payload = payload & """description"": { ""content"": [ { ""content"": [ { ""text"": " & ConvertToJson(description) & ", ""type"": ""text"" } ], ""type"": ""paragraph"" } ], ""type"": ""doc"", ""version"": 1 },"
           
        payload = payload & """issuetype"" : { ""name"" : """ & issueType & """ } } }"
        
        
    Set RestClient = RestClient.Service("POST", gstrAtlassianURL & "/rest/api/3/issue/", payload)
    
    If RestClient.status = 201 Then
        Dim json As Object
        Set json = JsonConverter.ParseJson(RestClient.responseText)
        CreateIssue = json("key")
    Else
        MsgBox RestClient.responseText, vbCritical, RestClient.status
    End If
       
End Function

Public Sub AddAttachment(key As String, filename As String, file As String)

    Dim payload() As Byte
    
    payload = StrConv("--" & gstrBoundary & vbCrLf & _
        "Content-Disposition: form-data; name=""file""; filename=""" & filename & """" & vbCrLf & _
        "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & StrConv(file, vbUnicode) & vbCrLf & _
        "--" & gstrBoundary & "--", vbFromUnicode)
        
    Call RestClient.Service("POST", gstrAtlassianURL & "/rest/api/3/issue/" & key & "/attachments", payload)
     
End Sub

Public Function GetAttachment(attachmentId As String) As Variant
    
    Set RestClient = RestClient.Service("GET", gstrAtlassianURL & "/rest/api/3/attachment/content/" & attachmentId)
    
    GetAttachment = RestClient.responseBody
    
End Function

Private Function getCloudId() As String
    
    getCloudId = gstrAtlassianCloudId
    
    If getCloudId = vbNullString Then
        Dim json As Object
        
        Set RestClient = RestClient.Service("GET", gstrAtlassianURL & "/_edge/tenant_info")
        Set json = JsonConverter.ParseJson(RestClient.responseText)
        getCloudId = json("cloudId")
    End If
    
End Function

Private Function AddUser(ByVal item As Object) As clsAtlassianUser

    Dim user As New clsAtlassianUser
    
    user.accountId = item("accountId")
    user.name = item("displayName")
    user.email = item("emailAddress")
    user.active = item("active")
        
    Set AddUser = user

End Function

Private Function AddIssueType(ByVal item As Object) As clsJiraIssuetype

    Dim issueType As New clsJiraIssuetype
    
    issueType.Id = item("id")
    issueType.name = item("name")
    issueType.subtask = item("subtask")
    
    Set AddIssueType = issueType

End Function

Private Function AddIssueProject(ByVal item As Object) As clsJiraProject

    Dim project As New clsJiraProject
    
    project.Id = item("id")
    project.key = item("key")
    project.name = item("name")
    
    Set AddIssueProject = project

End Function

Private Function AddIssueTransitions(ByVal items As Object) As Collection

    Dim transition As clsJiraIssueTransition
    Dim transitions As New Collection
    Dim item As Variant
    Dim subItem As Variant
    Dim tmpDate As Date
    Dim i As Integer
           
    'changed loop to get the first transitions first
    For i = items.Count To 1 Step -1
        Set item = items(i)
    
        For Each subItem In item("items")
                  
            If subItem("field") = "status" Then
                Set transition = New clsJiraIssueTransition
                
                If tmpDate = 0 Then tmpDate = ParseISO8601date(item("created"))
                
                transition.transitionDate = ParseISO8601date(item("created"))
                If Not IsNull(subItem("from")) Then transition.fromId = subItem("from")
                transition.fromString = subItem("fromString")
                transition.toId = subItem("to")
                transition.toString = subItem("toString")
                transition.daysInSourceStatus = DateDiff("d", tmpDate, ParseISO8601date(item("created")))
                transition.author = AddUser(item("author"))
                
                transitions.Add transition
                    
                tmpDate = ParseISO8601date(item("created"))
            End If
        
        Next
    Next i
    
    Set AddIssueTransitions = transitions
        
End Function

Private Function AddIssueComments(ByVal items As Object) As Collection

    Dim comment As clsJiraIssueComment
    Dim comments As New Collection
    Dim item As Variant
              
    For Each item In items
        Set comment = New clsJiraIssueComment
        
        comment.Id = item("id")
        comment.body = item("body")
        comment.created = ParseISO8601date(item("created"))
        
        'Check if author is anonymous
        If Not IsEmpty(item("author")) Then comment.author = AddUser(item("author"))
        
        comments.Add comment
    Next
        
    Set AddIssueComments = comments
        
End Function

Private Function AddAttachments(ByVal items As Object) As Collection

    Dim attachment As clsJiraIssueAttachment
    Dim attachments As New Collection
    Dim item As Variant
            
    For Each item In items
        Set attachment = New clsJiraIssueAttachment
        
        attachment.Id = item("id")
        attachment.filename = item("filename")
        
        attachments.Add attachment
    Next
        
    Set AddAttachments = attachments
        
End Function

Private Function AddIssueFixVersions(ByVal items As Object) As Collection

    Dim version As clsJiraIssueVersion
    Dim versions As New Collection
    Dim item As Variant
 
    For Each item In items
        Set version = New clsJiraIssueVersion
        
        version.Id = item("id")
        version.versionName = item("name")
        version.releaseDate = item("releaseDate")
        version.archived = item("archived")
        version.released = item("released")
        
        versions.Add version
    Next
        
    Set AddIssueFixVersions = versions
        
End Function

