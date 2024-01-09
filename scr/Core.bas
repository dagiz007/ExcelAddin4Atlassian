Attribute VB_Name = "Core"
'Open url in default browser
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public gstrAtlassianURL As String
Public gstrAtlassianEmail As String
Public gstrAtlassianToken As String
Public gblnLogging As Boolean
Public gstrLogPath As String

Public gblnSuccessfulLogin As Boolean
Public gstrBoundary As String

Public gobjJiraIssueCache As Object

Public gclsAppEvents As clsAppEvents                    'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
Public gcolAppEventResult As Collection

Sub Auto_Open()                                         'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in

    Set gclsAppEvents = New clsAppEvents
    Set gclsAppEvents.App = Application

    If gcolAppEventResult Is Nothing Then Set gcolAppEventResult = New Collection
    
    gstrAtlassianURL = GetSetting("ExcelAddin4Atlassian", "Settings", "AtlassianURL")
    gstrAtlassianEmail = GetSetting("ExcelAddin4Atlassian", "Settings", "AtlassianEmail")
    gstrAtlassianToken = GetSetting("ExcelAddin4Atlassian", "Settings", "AtlassianToken")
 
    'Logic to handle if settings are not set.
    gblnLogging = IIf(GetSetting("ExcelAddin4Atlassian", "Settings", "Logging") = "", False, GetSetting("ExcelAddin4Atlassian", "Settings", "Logging"))
 
    gstrLogPath = GetSetting("ExcelAddin4Atlassian", "Settings", "LogPath")
    
    If gstrAtlassianURL = vbNullString Then frmSettings.Show

End Sub

Public Sub openHyperlink(url)
    ShellExecute 0, vbNullString, url, vbNullString, vbNullString, vbNormalFocus
End Sub

'To open form from Outlook
Public Sub createJiraIssue()
    frmCreateJiraIssue.Show
End Sub

'To open form from Outlook
Public Sub openSettings()
    frmSettings.Show
End Sub

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

Public Function OpenExcelAddin4AtlassianSettings()
    
    Call addValueToResult
    
    frmSettings.Show
    
End Function

Public Function JiraOpenCreateIssueForm()
       
    Call addValueToResult
    
    frmCreateJiraIssue.Show
    
End Function

Public Function JiraCreateIssue(project As String, issueType As String, summary As String, description As String)
    
    Dim jiraKey As String
    jiraKey = Jira.CreateIssue(project, issueType, summary, description)
    
    Call addValueToResult(jiraKey)

End Function

Public Function JiraGetIssue(key As String)
    Call JiraGetIssues("key=" & key)
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

'Function to use clsConfluence class from other files.
Public Function Confluence() As clsConfluence
    Set Confluence = New clsConfluence
End Function

'Function to use clsJira class from other files.
Public Function Jira() As clsJira
    Set Jira = New clsJira
End Function

Public Function ReadFile(ByVal sFilepath As String) As Variant
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    
    With oStream
        .Type = 1
        .Open
        .LoadFromFile sFilepath
        ReadFile = .Read
        .Close
    End With
    
    Set oStream = Nothing
    
End Function

Public Sub WriteFile(filename As String, data As Variant)
    
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    
    With oStream
        .Open
        .Type = 1
        .Write data
        .SaveToFile (filename), 2
        .Close
    End With
        
    Set oStream = Nothing

End Sub

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

'TODO; check if we could ignore rest-response if binary.
Public Function IsASCII(str As String) As Boolean
    Dim i As Long
    For i = 1 To Len(str)
        If AscW(Mid(str, i, 1)) > 127 Then
            IsASCII = False
            Exit Function
        End If
    Next i
    IsASCII = True
End Function

Public Sub addLog(filename As String, text As String)
    
    On Error GoTo Errorhandler
    
    filename = gstrLogPath & "\" & filename

    Dim strFile_Path As String
    Dim objFSO As Object
    Dim objFile As Object
    
    Dim fileSize As Long
    Dim fileExtensionSeparator As Integer
    Dim fileSeparator As Integer
        
    'TODO; refactors
    Do
            If Len(Dir(filename)) = 0 Then
                fileSize = 0
            Else
                fileSize = FileLen(filename)
            End If
    
            If fileSize >= 500000000 Then '500MB
                fileExtensionSeparator = InStrRev(filename, ".")
                
                fileSeparator = InStrRev(filename, "_")
                
                If fileExtensionSeparator = 0 Then
                    If Not IsNumeric(Mid(filename, fileSeparator + 1)) Then fileSeparator = 0
                Else
                    If Not IsNumeric(Mid(filename, fileSeparator + 1, fileExtensionSeparator - fileSeparator - 1)) Then fileSeparator = 0
                End If
                
                
                If fileSeparator = 0 Then
                    If fileExtensionSeparator = 0 Then
                        filename = filename & "_0"
                    Else
                        filename = Left(filename, fileExtensionSeparator - 1) & "_0" & Mid(filename, fileExtensionSeparator)
                    End If
                Else
                    If fileExtensionSeparator = 0 Then
                        filename = Left(filename, fileSeparator) & Mid(filename, fileSeparator + 1) + 1
                    Else
                        filename = Left(filename, fileSeparator) & Mid(filename, fileSeparator + 1, fileExtensionSeparator - fileSeparator - 1) + 1 & Mid(filename, fileExtensionSeparator)
                    End If
                End If
                
            Else
                Exit Do
            End If
    Loop
    
        
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(filename, 8, True, -1)
    
    objFile.WriteLine Format(Now(), "YYYY-MM-DD HH:MM:SS") & vbTab & text
    objFile.Close
    
    Exit Sub

Errorhandler:

    Debug.Print Err.Number & " " & Err.description
    Stop
    Exit Sub

End Sub
