Attribute VB_Name = "CoreExcel"
Public gclsAppEvents As clsAppEvents                    'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
Public gcolAppEventResult As Collection

Sub Auto_Open()                                         'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in

    Set gclsAppEvents = New clsAppEvents
    Set gclsAppEvents.App = Application

    If gcolAppEventResult Is Nothing Then Set gcolAppEventResult = New Collection
    
    Call LoadSettings
  
End Sub

Public Function OpenExcelAddin4AtlassianSettings()
    
    Call addValueToResult
    
    frmSettings.Show
    
End Function

Public Function JiraCreateIssue(project As String, issueType As String, summary As String, description As String)
    
    Dim jiraKey As String
    jiraKey = Jira.CreateIssue(project, issueType, summary, description)
    
    Call addValueToResult(jiraKey)

End Function

Public Function JiraDownloadIssusAttachments(jql As String, path As String)
    
    frmWait.Show vbModeless
       
    If Right(path, 1) = "\" Then path = Left(path, Len(path) - 1)
    If Dir(path, vbDirectory) = Empty Then MkDir path
    
    Dim issues As Collection
    Dim issue As clsJiraIssue
    
    Dim attachment As clsJiraIssueAttachment
    Dim counter As Integer
        
    Set issues = Jira.GetIssues(jql)
    
    For Each issue In issues
        counter = 1
        For Each attachment In issue.attachment
            Call WriteFile(path & "\" & issue.key & "_" & counter & "_" & attachment.filename, Jira.GetAttachment(attachment.id))
            counter = counter + 1
        Next
    Next
    
    Call addValueToResult("Attachments are downloaded to " & path)
       
End Function

Public Function JiraGetIssue(key As String)
    Call JiraGetIssues("key=" & key)
End Function

Public Function JiraGetIssueDaysInTransitions(jiraKey As String, ParamArray transitions() As Variant) As Integer

    Dim transition As clsJiraIssueTransition
    Dim issue As clsJiraIssue
    
    Set issue = Jira.GetIssue(jiraKey)
    
    For Each transition In issue.transition
        If IsInArray(transition.fromString, CVar(transitions)) Then
            JiraGetIssueDaysInTransitions = JiraGetIssueDaysInTransitions + transition.daysInSourceStatus
        End If
    Next
    
End Function

Public Function JiraGetIssues(jql As String)
        
    frmWait.Show vbModeless

    Dim issues As Collection
    Dim issue As clsJiraIssue
    
    Set issues = Jira.GetIssues(jql)
    
    Dim row As Integer
    
    Dim bdTable As New clsBreakDownTable
    bdTable.startingPosition = Range(ActiveCell.Address)
    
    For Each issue In issues
        Call addValueToResult(issue.key, bdTable.GetCellPosition(row))
        Call addValueToResult(issue.summary, bdTable.GetCellPosition(row, 1))
        row = row + 1
    Next
    
End Function

Public Function JiraOpenCreateIssueForm()
       
    Call addValueToResult
    
    frmCreateJiraIssue.Show
    
End Function

Private Sub addValueToResult(Optional ByVal cellvalue As String = "", Optional position As Range, Optional cellformat As String = "General")
 
    'Set default value
    If position Is Nothing Then Set position = Range(ActiveCell.Address)
    
    If gclsAppEvents Is Nothing Then Call Auto_Open
    
    Dim bdTable As New clsBreakDownTable

    bdTable.cellvalue = cellvalue
    bdTable.startingPosition = position
    bdTable.cellformat = cellformat
    
    gcolAppEventResult.Add bdTable
   
End Sub
