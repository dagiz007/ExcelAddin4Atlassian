VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJiraIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pId As String
Private pKey As String
Private pSummary As String
Private pReporter As clsJiraUser
Private pAssignee As clsJiraUser
Private pCreator As clsJiraUser
Private pIssueType As clsJiraIssuetype
Private pTransition As Collection
Private pVersion As Collection
Private pProject As clsJiraProject
Private pCreatedDate As Date
Private pResolutionDate As Date
Private pAttachment As Collection
Private pComment As Collection
Private pDescription As String
Private pDescriptionRendered As String
Private pJson As Object

Public Property Get Id() As String
    Id = pId
End Property

Public Property Let Id(value As String)
    pId = value
End Property

Public Property Get key() As String
    key = pKey
End Property

Public Property Let key(value As String)
    pKey = value
End Property

Public Property Get summary() As String
    summary = pSummary
End Property

Public Property Let summary(value As String)
    pSummary = value
End Property

Public Property Get assignee() As clsJiraUser
    Set assignee = pAssignee
End Property

Public Property Let assignee(value As clsJiraUser)
    Set pAssignee = value
End Property

Public Property Get reporter() As clsJiraUser
    Set reporter = pReporter
End Property

Public Property Let reporter(value As clsJiraUser)
    Set pReporter = value
End Property

Public Property Get creator() As clsJiraUser
    Set creator = pCreator
End Property

Public Property Let creator(value As clsJiraUser)
    Set pCreator = value
End Property

Public Property Get issueType() As clsJiraIssuetype
    Set issueType = pIssueType
End Property

Public Property Let issueType(value As clsJiraIssuetype)
    Set pIssueType = value
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

Public Property Get project() As clsJiraProject
    Set project = pProject
End Property

Public Property Let project(value As clsJiraProject)
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

Public Property Get description() As String
     description = pDescription
End Property

Public Property Let description(value As String)
     pDescription = value
End Property

Public Property Get descriptionRendered() As String
     descriptionRendered = pDescriptionRendered
End Property

Public Property Let descriptionRendered(value As String)
     pDescriptionRendered = value
End Property

Public Property Get json() As Object
     Set json = pJson
End Property

Public Property Let json(value As Object)
     Set pJson = value
End Property


