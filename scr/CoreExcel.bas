Attribute VB_Name = "CoreExcel"
Public gclsAppEvents As clsAppEvents                    'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
Public gcolAppEventResult As Collection

Sub Auto_Open()                                         'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in

    Set gclsAppEvents = New clsAppEvents
    Set gclsAppEvents.App = Application

    If gcolAppEventResult Is Nothing Then Set gcolAppEventResult = New Collection
    
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
            Call WriteFile(path & "\" & issue.key & "_" & counter & "_" & attachment.filename, Jira.GetAttachment(attachment.Id))
            counter = counter + 1
        Next
    Next
    
    Call addValueToResult("Attachments are downloaded to " & path)
       
End Function

Public Function JiraGetIssueFieldValue(key As String, value As String) As String
    Dim issue As clsJiraIssue
    Set issue = Jira.GetIssue(key)
      
    JiraGetIssueFieldValue = issue.json("fields")("" & value & "")
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

Public Function AtlassianGetAllUsers()
        
    frmWait.Show vbModeless
    Dim users As Collection
    Dim user As clsAtlassianUser
    Dim product As clsAtlassianProductAccess
        
    Set users = Atlassian.GetUsers
        
    Dim row As Integer
    
    Dim bdTable As New clsBreakDownTable
    bdTable.startingPosition = Range(ActiveCell.Address)
    
    Call addValueToResult("Name", bdTable.GetCellPosition(row))
    Call addValueToResult("Email", bdTable.GetCellPosition(row, 1))
    Call addValueToResult("Active", bdTable.GetCellPosition(row, 2))
    Call addValueToResult("Product", bdTable.GetCellPosition(row, 3))
    Call addValueToResult("URL", bdTable.GetCellPosition(row, 4))
    Call addValueToResult("Last active", bdTable.GetCellPosition(row, 5))
    row = row + 1
    
    For Each user In users
        For Each product In user.productAccess
            Call addValueToResult(user.name, bdTable.GetCellPosition(row))
            Call addValueToResult(user.email, bdTable.GetCellPosition(row, 1))
            Call addValueToResult(user.active, bdTable.GetCellPosition(row, 2))
            Call addValueToResult(product.name, bdTable.GetCellPosition(row, 3))
            Call addValueToResult(product.url, bdTable.GetCellPosition(row, 4))
            If product.lastActive <> 0 Then Call addValueToResult(product.lastActive, bdTable.GetCellPosition(row, 5), "dd.MM.yyyy HH:mm:ss")
            row = row + 1
        Next
    Next
    
End Function

Public Function JiraGetFirstFormId(key As String)
    JiraGetFirstFormId = Jira.GetFormId(key)(1).Id
End Function
