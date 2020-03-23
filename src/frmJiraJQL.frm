VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJiraJQL 
   Caption         =   "Jira JQL"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8895.001
   OleObjectBlob   =   "frmJiraJQL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmJiraJql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim jiraClient As New JiraRestClient
    Dim issues As Collection
    Dim issue As issue

    Set issues = jiraClient.getJiraIssues(Trim(txtJQL))
      
    Dim row As Integer
    row = 0
    
    Dim bdTable As New BreakDownTable
    bdTable.StartingPosition = Range(refEditTabel)
          
    For Each issue In issues
        bdTable.Cell(row, 0).NumberFormat = "General"
        bdTable.Cell(row, 0) = issue.jiraKey
        row = row + 1
    Next

    SaveSetting "ExcelAddIn4Jira", "JiraJQL", "Jql", Trim(txtJQL)
    Unload Me
    
End Sub

Private Sub lblExcelAddin4Jira_Click()
    ActiveWorkbook.FollowHyperlink "https://github.com/DagAtleStenstad/ExcelAdd-in4Jira"
End Sub

Private Sub UserForm_Initialize()
    refEditTabel = activeCellAddress
    txtJQL = GetSetting("ExcelAddIn4Jira", "JiraJQL", "Jql")
End Sub
