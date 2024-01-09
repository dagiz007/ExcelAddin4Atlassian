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

Public Sub LoadSettings()
    
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
Public Sub OpenCreateJiraIssueForm()
    If gstrAtlassianURL = vbNullString Then Call LoadSettings
    frmCreateJiraIssue.Show
End Sub

'To open form from Outlook
Public Sub OpenSettings()
    If gstrAtlassianURL = vbNullString Then Call LoadSettings
    frmSettings.Show
End Sub

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
