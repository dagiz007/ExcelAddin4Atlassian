VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JiraRestClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private jiraUrl As String
Private jiraUsr As String
Private jiraPwd As String

Private Sub Class_initialize()

    If jiraUrl = vbNullString Or jiraUsr = vbNullString Or jiraPwd = vbNullString Then
        jiraUrl = GetSetting("ExcelAddIn4Jira", "Settings", "Jira_url")
        jiraUsr = GetSetting("ExcelAddIn4Jira", "Settings", "Jira_username")
        jiraPwd = GetSetting("ExcelAddIn4Jira", "Settings", "Jira_password")
    End If
      
    If jiraUrl = vbNullString Or jiraUsr = vbNullString Or jiraPwd = vbNullString Then
        frmJiraSettings.Show
    End If
    
End Sub

Private Function JiraService(methods As String, url As String, Optional sData As Variant, Optional responseAsBytes As Boolean) As Variant
    
    Dim jiraClient As Object
    Dim jiraResponse As Variant

    Set jiraClient = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With jiraClient
        .Open methods, jiraUrl & url, False
        .setRequestHeader "Authorization", "Basic " & EncodeBase64(jiraUsr & ":" & jiraPwd)
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .Send sData
        
        If responseAsBytes Then
            jiraResponse = .responseBody
        Else
            jiraResponse = .ResponseText
        End If
       
    End With
        
    Set jiraClient = Nothing
    
    JiraService = jiraResponse
    
End Function

Public Sub saveJiraAttachmentToFile(attachmentId, saveAsFilename)

    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    
    Dim jiraResponse As Variant
    jiraResponse = JiraService("GET", "secure/attachment/" & attachmentId & "/", , True)
    
    With oStream
        .Open
        .Type = 1
        .Write jiraResponse
        .SaveToFile (saveAsFilename), 2
        .Close
    End With
        
    Set oStream = Nothing

End Sub

Public Function getJiraStatus() As Collection

    Dim jiraResponse As String
    Dim json As Object
    Dim item As Variant
    Dim status As status
    Dim statuses As New Collection
    
    jiraResponse = JiraService("GET", "rest/api/2/status")
    
    Set json = JsonConverter.ParseJson(jiraResponse)
    
    For Each item In json
        Set status = New status
        
        status.id = item("id")
        status.name = item("name")
        
        statuses.Add status
    Next
    
    Set getJiraStatus = statuses

End Function

Public Function getJiraProject() As Collection

    Dim jiraResponse As String
    Dim json As Object
    Dim item As Variant
    Dim project As project
    Dim projects As New Collection
    
    jiraResponse = JiraService("GET", "rest/api/2/project")
    
    Set json = JsonConverter.ParseJson(jiraResponse)
    
    For Each item In json
        Set project = New project
        
        project.id = item("id")
        project.key = item("key")
        project.name = item("name")
        
        projects.Add project
    Next
    
    Set getJiraProject = projects

End Function

Public Function getIssueTypes(jiraProjectID As String) As Collection
    
    Dim jiraResponse As String
    Dim json As Object
    Dim item As Variant
    Dim issueType As issueType
    Dim issueTypes As New Collection
        
    jiraResponse = JiraService("GET", "rest/api/2/project/" & jiraProjectID & "/statuses")
    
    Set json = JsonConverter.ParseJson(jiraResponse)
    
     For Each item In json
        Set issueType = New issueType
            
        issueType.id = item("id")
        issueType.name = item("name")
        issueType.subtask = item("subtask")
        
        issueTypes.Add issueType
    Next
    
    Set getIssueTypes = issueTypes
    
End Function

Public Function getJiraConfiguration() As Variant
    
    Dim jiraResponse As String
    Dim json As Object
    Dim configuration(2) As Variant

    jiraResponse = JiraService("GET", "rest/api/2/configuration")
    
    Set json = JsonConverter.ParseJson(jiraResponse)
    
    configuration(0) = json("timeTrackingConfiguration")("workingHoursPerDay")
    configuration(1) = json("timeTrackingConfiguration")("workingDaysPerWeek")
           
    getJiraConfiguration = configuration
    
End Function

Public Function getJiraIssues(jql As String) As Collection

    Dim data As String
    Dim jiraResponse As String
    Dim json As Object
    Dim item As Variant
    Dim issue As issue
    Dim issues As New Collection
    Dim maxResults As Integer
    Dim startAt As Integer
    Dim totalIssues As Integer
    
    startAt = 0
    maxResults = 50
    
    jql = Replace(jql, """", "'")

readIssues:
    
    data = " { ""jql"": """ & Trim(jql) & """, ""startAt"": " & startAt & ",""maxResults"": " & maxResults & _
           ",""fields"": [], ""expand"": [] } "
       
    jiraResponse = JiraService("POST", "rest/api/2/search", data)

    Set json = JsonConverter.ParseJson(jiraResponse)
    
    totalIssues = json("total")
    
    For Each item In json("issues")
        Set issue = New issue
        
        Set issue = getJiraIssue(item("key"))
        
        issues.Add issue

    Next
    
    startAt = startAt + maxResults
    If startAt < totalIssues Then GoTo readIssues
    
    Set getJiraIssues = issues
    
End Function

Public Function getJiraIssue(ByVal jiraKey As String) As issue

    Dim json As Object
    Dim issue As New issue
    Dim project As New project
    Dim issueType As New issueType
    Dim data As String
    
    '//Cache JiraIssueResponse
    If jiraIssueCache Is Nothing Then Set jiraIssueCache = CreateObject("Scripting.Dictionary")
    
    If Not jiraIssueCache.Exists(jiraKey) Then
        jiraIssueCache.Add jiraKey, JiraService("GET", "rest/api/2/issue/" & Trim(jiraKey) & "?expand=changelog")
    End If

    Set json = JsonConverter.ParseJson(jiraIssueCache(jiraKey))
    
    issue.id = json("id")
    issue.jiraKey = json("key")
    issue.summary = json("fields")("summary")
    
    'TO DO, add assagee and reporter as own objects.
    issue.assignee = json("fields")("assignee")("displayName")
    issue.reporter = json("fields")("reporter")("displayName")
    
    issueType.id = json("fields")("issuetype")("id")
    issueType.name = json("fields")("issuetype")("name")
    issueType.subtask = json("fields")("issuetype")("subtask")
    issue.issueType = issueType
    
    issue.originalEstimateHours = IIf(IsNull(json("fields")("timeoriginalestimate")), "", json("fields")("timeoriginalestimate"))
    issue.remainingEstimateHours = IIf(IsNull(json("fields")("timeestimate")), "", json("fields")("timeestimate"))
    issue.originalAggregateEstimateHours = IIf(IsNull(json("fields")("aggregatetimeoriginalestimate")), "", json("fields")("aggregatetimeoriginalestimate"))
    issue.remainingAggregateEstimateHours = IIf(IsNull(json("fields")("aggregatetimeestimate")), "", json("fields")("aggregatetimeestimate"))
    issue.aggregateProgress = IIf(IsNull(json("fields")("aggregateprogress")("progress")), "", json("fields")("aggregateprogress")("progress"))
    
    issue.worklog = AddIssueWorklogs(json("fields")("worklog")("worklogs"))
    issue.transition = AddIssueTransitions(json("changelog")("histories"))
    issue.attachment = AddAttachments(json("fields")("attachment"))
    issue.version = AddIssueFixVersions(json("fields")("fixVersions"))
    issue.createdDate = parseISO8601date(json("fields")("created"))
    
    If Not IsNull(json("fields")("resolutiondate")) Then
        issue.resolutionDate = parseISO8601date(json("fields")("resolutiondate"))
    End If
    
    project.id = json("fields")("project")("id")
    project.key = json("fields")("project")("key")
    project.name = json("fields")("project")("name")
    issue.project = project
    
    Set getJiraIssue = issue
    
End Function

Private Function AddIssueFixVersions(ByVal items As Object) As Collection

    Dim version As version
    Dim versions As New Collection
    Dim item As Variant
 
    For Each item In items
        Set version = New version
        
        version.id = item("id")
        version.versionName = item("name")
        version.releaseDate = item("releaseDate")
        version.archived = item("archived")
        version.released = item("released")
        
        versions.Add version
    Next
        
    Set AddIssueFixVersions = versions
        
End Function

Private Function AddIssueWorklogs(ByVal items As Object) As Collection

    Dim worklog As worklog
    Dim worklogs As New Collection
    Dim item As Variant
            
    For Each item In items
        Set worklog = New worklog
        
        worklog.id = item("id")
        worklog.timeSpentSeconds = item("timeSpentSeconds")
        worklog.started = parseISO8601date(item("started"))
        
        worklogs.Add worklog
    Next
        
    Set AddIssueWorklogs = worklogs
        
End Function

Private Function AddAttachments(ByVal items As Object) As Collection

    Dim attachment As attachment
    Dim attachments As New Collection
    Dim item As Variant
            
    For Each item In items
        Set attachment = New attachment
        
        attachment.id = item("id")
        attachment.fileName = item("filename")
        
        attachments.Add attachment
    Next
        
    Set AddAttachments = attachments
        
End Function

Private Function AddIssueTransitions(ByVal items As Object) As Collection

    Dim transition As transition
    Dim transitions As New Collection
    Dim item As Variant
    Dim subItem As Variant
    Dim tmpDate As Date
           
    For Each item In items
        For Each subItem In item("items")
                  
            If tmpDate = 0 Then tmpDate = parseISO8601date(item("created"))
            
            If subItem("field") = "status" Then
                Set transition = New transition
                
                transition.transitionDate = parseISO8601date(item("created"))
                transition.fromId = subItem("from")
                transition.fromString = subItem("fromString")
                transition.toId = subItem("to")
                transition.toString = subItem("toString")
                transition.timeInSourceStatus = DateDiff("d", tmpDate, parseISO8601date(item("created")))
                
                transitions.Add transition
                    
                tmpDate = parseISO8601date(item("created"))
            End If
        
        Next
    Next
    
    Set AddIssueTransitions = transitions
        
End Function

Public Sub DeleteJiraIssue(sIssueID As String, bDeleteSubTasks As Boolean)
    Dim jiraResponse As String
    jiraResponse = JiraService("DELETE", "rest/api/2/issue/" & sIssueID & "?deleteSubtasks=" & bDeleteSubTasks)
    
     'Todo, check if status is sucsedsfull
End Sub

Public Function addWatcher(jiraKey As String, watcher As String) As Boolean
    Dim jiraUser As Variant
    jiraUser = getJiraUserKey(watcher)
    
    If Not IsNull(jiraUser) Then
        watcher = """" & jiraUser & """"
        Call JiraService("POST", "rest/api/2/issue/" & jiraKey & "/watchers", watcher)
        addWatcher = True
    Else
        addWatcher = False
    End If
End Function

Public Function getUsers(fuser As String) As Collection
    
    Dim jiraResponse As String
    Dim json As Object
    Dim item As Variant
    Dim users As New Collection
    Dim user As user

    jiraResponse = JiraService("GET", "rest/api/2/user/search?username=" & fuser)
    
    Set json = JsonConverter.ParseJson(jiraResponse)
    
    For Each item In json
        Set user = New user
    
            user.userKey = item("key")
            user.email = item("emailAddress")
            user.name = item("displayName")
            user.active = item("active")
            
            users.Add user
    Next
       
    Set getUsers = users
    
End Function

Public Function getJiraUserKey(findUser As String) As Variant

    Dim users As Collection
    Dim user As user
    Set users = getUsers(findUser)
    
    If users.Count = 1 Then
        Set user = users.item(1)
        getJiraUserKey = user.userKey
    Else
        getJiraUserKey = Null
    End If

End Function

Public Sub JiraLogout()
    Call JiraService("DELETE", "rest/auth/1/session")
End Sub

Private Function parseISO8601date(ByVal iso8601date As String) As Date
    parseISO8601date = Mid(iso8601date, 9, 2) & "." & Mid(iso8601date, 6, 2) & "." & Left(iso8601date, 4) & " " & Mid(iso8601date, 12, 8)
End Function

Private Function EncodeBase64(text As String) As String
    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
  
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
  
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
  
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text
  
    Set objNode = Nothing
    Set objXML = Nothing
End Function
