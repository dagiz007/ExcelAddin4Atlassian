Attribute VB_Name = "Core"
Option Explicit

Public activeCellAddress As String
Public excelAddInn4JiraCommand As String
Public excelAddInn4JiraCommandArg As String
Public jiraIssueCache As Object

Private jiraClient As New JiraRestClient
Private issue As issue

Public gclsAppEvents As ExcelAddIn4JiraAppEvents        'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
  
Sub Auto_Open()                                         'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
    Set gclsAppEvents = New ExcelAddIn4JiraAppEvents
    Set gclsAppEvents.App = Application
End Sub

Function JiraDownloadIssuesAttachments() As String
    If gclsAppEvents Is Nothing Then Call Auto_Open
    
    activeCellAddress = ActiveCell.Address
    excelAddInn4JiraCommand = "openJiraDownloadIssusAttachmentsForm"
End Function

Function JiraGetIssues() As String
   If gclsAppEvents Is Nothing Then Call Auto_Open
   
   activeCellAddress = ActiveCell.Address
   excelAddInn4JiraCommand = "openJiraJQLform"
End Function

Function JiraSettings() As String
    If gclsAppEvents Is Nothing Then Call Auto_Open
    
    activeCellAddress = ActiveCell.Address
    excelAddInn4JiraCommand = "openJiraSettingsForm"
End Function

Function JiraGetIssueSummary(jiraKey As String) As String
    Set issue = jiraClient.getJiraIssue(jiraKey)
    JiraGetIssueSummary = issue.summary
End Function

Function JiraGetIssueCreatedDate(jiraKey As String) As Date
    Set issue = jiraClient.getJiraIssue(jiraKey)
    JiraGetIssueCreatedDate = issue.createdDate
End Function

Function JiraGetIssueIssueType(jiraKey As String) As String
    Set issue = jiraClient.getJiraIssue(jiraKey)
    JiraGetIssueIssueType = issue.issueType.name
End Function

Function JiraGetIssueAssignee(jiraKey As String) As String
    Set issue = jiraClient.getJiraIssue(jiraKey)
    JiraGetIssueAssignee = issue.assignee
End Function

Function JiraGetIssueReporter(jiraKey As String) As String
    Set issue = jiraClient.getJiraIssue(jiraKey)
    JiraGetIssueReporter = issue.reporter
End Function

Function JiraGetIssueDaysInTransitions(jiraKey As String, ParamArray transitions() As Variant) As Integer

Dim transition As transition
Dim status As String

Set issue = jiraClient.getJiraIssue(jiraKey)
    
    For Each transition In issue.transition
        
        status = transition.fromString
        
        If IsInArray(status, CVar(transitions)) Then
            JiraGetIssueDaysInTransitions = JiraGetIssueDaysInTransitions + transition.timeInSourceStatus
        End If
        
    Next
    
End Function

Function JiraGetIssueLatestReleaseDate(jiraKey As String) As Date
    Set issue = jiraClient.getJiraIssue(jiraKey)
    Dim version As version
       
    JiraGetIssueLatestReleaseDate = "00:00:00"
            
    For Each version In issue.version
        If Not version.releaseDate = "00:00:00" Then
            If JiraGetIssueLatestReleaseDate = "00:00:00" Or JiraGetIssueLatestReleaseDate <= version.releaseDate Then
                JiraGetIssueLatestReleaseDate = version.releaseDate
            End If
        End If
    Next
    
    'TO Do, set date if not resolutionDate is set.
    If JiraGetIssueLatestReleaseDate = "00:00:00" Then
        JiraGetIssueLatestReleaseDate = Format(issue.resolutionDate, "Short Date")
    End If
        
    JiraGetIssueLatestReleaseDate = JiraGetIssueLatestReleaseDate
End Function

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
