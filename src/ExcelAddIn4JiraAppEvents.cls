VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelAddIn4JiraAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents mxlApp As Application       'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
Attribute mxlApp.VB_VarHelpID = -1

Public Property Set App(xlApp As Application)  'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
    Set mxlApp = xlApp
End Property

Private Sub mxlApp_SheetCalculate(ByVal Sh As Object)
    
    If Not excelAddInn4JiraCommand = "" Then
    
        Select Case excelAddInn4JiraCommand
            
            Case "openJiraJQLform"
                 Range(activeCellAddress) = ""
                 Range(activeCellAddress).Select
                 frmJiraJql.Show
                
            Case "openJiraSettingsForm"
                 Range(activeCellAddress) = ""
                 Range(activeCellAddress).Select
                 frmJiraSettings.Show
                 
            Case "openJiraDownloadIssusAttachmentsForm"
                 Range(activeCellAddress) = ""
                 Range(activeCellAddress).Select
                 frmJiraDownloaIssusAttachments.Show
                 
        End Select
    
    End If
 
    excelAddInn4JiraCommand = ""
    
End Sub

