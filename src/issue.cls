VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private psJiraKey As String
Private pSummary As String
Private pReporter As user
Private pAssignee As user
Private pCreator As user
Private pIssueType As issueType
Private pOriginalEstimateHours As String
Private pRemainingEstimateHours As String
Private pAggregateProgress As String
Private pOriginalAggregateEstimateHours As String
Private pRemainingAggregateEstimateHours As String
Private pWorklog As Collection
Private pTransition As Collection
Private pVersion As Collection
Private pProject As project
Private pCreatedDate As Date
Private pResolutionDate As Date
Private pAttachment As Collection
Private pComment As Collection
Private pJson As Object

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(value As String)
    pId = value
End Property

Public Property Get jiraKey() As String
    jiraKey = psJiraKey
End Property

Public Property Let jiraKey(value As String)
    psJiraKey = value
End Property

Public Property Get summary() As String
    summary = pSummary
End Property

Public Property Let summary(value As String)
    pSummary = value
End Property

Public Property Get assignee() As user
    Set assignee = pAssignee
End Property

Public Property Let assignee(value As user)
    Set pAssignee = value
End Property

Public Property Get reporter() As user
    Set reporter = pReporter
End Property

Public Property Let reporter(value As user)
    Set pReporter = value
End Property

Public Property Get creator() As user
    Set creator = pCreator
End Property

Public Property Let creator(value As user)
    Set pCreator = value
End Property

Public Property Get issueType() As issueType
    Set issueType = pIssueType
End Property

Public Property Let issueType(value As issueType)
    Set pIssueType = value
End Property

Public Property Get originalEstimateHours() As String
    originalEstimateHours = pOriginalEstimateHours
End Property

Public Property Let originalEstimateHours(value As String)
    If Not value = "" Then
        pOriginalEstimateHours = value / 3600
    Else
        pOriginalEstimateHours = "0"
    End If
End Property

Public Property Get remainingEstimateHours() As String
    remainingEstimateHours = pRemainingEstimateHours
End Property

Public Property Let remainingEstimateHours(value As String)
    If Not value = "" Then
        pRemainingEstimateHours = value / 3600
    Else
        pRemainingEstimateHours = "0"
    End If
End Property

Public Property Get aggregateProgress() As String
    aggregateProgress = pAggregateProgress
End Property

Public Property Let aggregateProgress(value As String)
    If Not value = "" Then
        pAggregateProgress = value / 3600
    Else
        pAggregateProgress = "0"
    End If
End Property

Public Property Get originalAggregateEstimateHours() As String
    originalAggregateEstimateHours = pOriginalAggregateEstimateHours
End Property

Public Property Let originalAggregateEstimateHours(value As String)
    If Not value = "" Then
        pOriginalAggregateEstimateHours = value / 3600
    Else
        pOriginalAggregateEstimateHours = "0"
    End If
End Property

Public Property Get remainingAggregateEstimateHours() As String
    remainingAggregateEstimateHours = pRemainingAggregateEstimateHours
End Property

Public Property Let remainingAggregateEstimateHours(value As String)
    If Not value = "" Then
        pRemainingAggregateEstimateHours = value / 3600
    Else
        pRemainingAggregateEstimateHours = "0"
    End If
End Property

Public Property Get worklog() As Collection
    Set worklog = pWorklog
End Property

Public Property Let worklog(value As Collection)
    Set pWorklog = value
End Property

Public Property Get transition() As Collection
    Set transition = pTransition
End Property

Public Property Let transition(value As Collection)
    Set pTransition = value
End Property

Public Property Get version() As Collection
    Set version = pVersion
End Property

Public Property Let version(value As Collection)
    Set pVersion = value
End Property

Public Property Get project() As project
    Set project = pProject
End Property

Public Property Let project(value As project)
    Set pProject = value
End Property

Public Property Get createdDate() As Date
    createdDate = pCreatedDate
End Property

Public Property Let createdDate(value As Date)
    pCreatedDate = value
End Property

Public Property Get resolutionDate() As Date
    resolutionDate = pResolutionDate
End Property

Public Property Let resolutionDate(value As Date)
    pResolutionDate = value
End Property

Public Property Get attachment() As Collection
    Set attachment = pAttachment
End Property

Public Property Let attachment(value As Collection)
    Set pAttachment = value
End Property

Public Property Get comment() As Collection
    Set comment = pComment
End Property

Public Property Let comment(value As Collection)
    Set pComment = value
End Property

Public Property Get json() As Object
    Set json = pJson
End Property

Public Property Let json(value As Object)
    Set pJson = value
End Property



