VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfluence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private RestClient As New clsRestClient

Private Sub Class_initialize()
    If Not gblnSuccessfulLogin Then
        If Not CorrectCredentianls Then
            frmSettings.Show
        Else
            gblnSuccessfulLogin = True
        End If
    End If
End Sub

Private Sub Class_Terminate()
    Set RestClient = Nothing
End Sub

Public Function CorrectCredentianls() As Boolean
    If Jira.CorrectCredentianls Then CorrectCredentianls = True
End Function

Public Function GetSpaces() As Collection
    
    On Error GoTo Errorhandler

    Dim url As String
    
    Dim json As Object
    Dim item As Variant
    
    Dim space As clsConfluenceSpace
    Dim spaces As New Collection
    
    url = "/wiki/api/v2/spaces"
        
    Do
        DoEvents
        
        Set RestClient = RestClient.Service("GET", url)
        Set json = JsonConverter.ParseJson(RestClient.responseText)
        
        For Each item In json("results")
            Set space = New clsConfluenceSpace
            
            space.id = item("id")
            space.name = item("name")
            space.key = item("key")
            space.spaceType = item("type")
            
            spaces.Add space
        Next
        
        url = json("_links")("next")
        
    Loop Until url = ""
    
    Set GetSpaces = spaces
    
    Exit Function
    
    
Errorhandler:

    Debug.Print Err.Number & Err.description
    Stop
    Resume Next
    
End Function

Public Function GetPages(Optional spaceId As Long) As Collection

    On Error GoTo Errorhandler

    Dim url As String
    
    Dim json As Object
    Dim item As Variant
    
    Dim page As clsConfluencePage
    Dim pages As New Collection
    
    If spaceId = 0 Then
        url = "/wiki/api/v2/pages?body-format=storage"
    Else
        url = "/wiki/api/v2/spaces/" & spaceId & "/pages?body-format=storage"
    End If
        
    Do
        DoEvents
        
        Set RestClient = RestClient.Service("GET", url)
        Set json = JsonConverter.ParseJson(RestClient.responseText)
        
        For Each item In json("results")
            Set page = New clsConfluencePage
            Set page = GetPage(item("id"), item)
            pages.Add page
        Next
        
        url = json("_links")("next")
        
    Loop Until url = ""
    
    Set GetPages = pages
    
    Exit Function
    
    
Errorhandler:

    Debug.Print Err.Number & Err.description
    Stop
    Resume Next
    
End Function

Public Function GetPage(ByVal id As String, Optional ByVal json As Object) As clsConfluencePage

    On Error GoTo Errorhandler
        
    If json Is Nothing Then
        Set RestClient = RestClient.Service("get", "/wiki/api/v2/pages/" & id & "?body-format=storage")
        Set json = JsonConverter.ParseJson(RestClient.responseText)
    End If
    
    Dim page As New clsConfluencePage
                
    page.id = json("id")
    page.title = json("title")
    page.version = json("version")("number")
    page.body = json("body")("storage")("value")
    page.spaceId = json("spaceId")
    page.status = json("status")
        
    Set GetPage = page

    Exit Function
    
    
Errorhandler:

    Debug.Print Err.Number & Err.description
    Stop
    Resume Next

End Function

Public Function UpdateConfluenceContent(id As String, status As String, title As String, body As String, version As Integer) As clsRestClient

   Dim payload As String

   payload = "{""id"":""" & id & """," & _
             " ""status"":""" & status & """," & _
             " ""title"":""" & title & """," & _
             " ""body"":{""representation"":""storage"", ""value"":" & JsonConverter.ConvertToJson(body) & "}," & _
             " ""version"":{""number"":" & version & "}}"
    
   Set UpdateConfluenceContent = RestClient.Service("put", "/wiki/api/v2/pages/" & id, payload)
    
End Function
