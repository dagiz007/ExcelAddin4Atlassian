VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJiraIssueComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pCreated As Date
Private pId As Long
Private pBody As String
Private pAuthor As clsJiraUser

Public Property Get created() As Date
    created = pCreated
End Property

Public Property Let created(value As Date)
    pCreated = value
End Property

Public Property Get id() As Long
    id = pId
End Property

Public Property Let id(value As Long)
    pId = value
End Property

Public Property Get body() As String
    body = pBody
End Property

Public Property Let body(value As String)
    pBody = value
End Property

Public Property Get author() As clsJiraUser
    Set author = pAuthor
End Property

Public Property Let author(value As clsJiraUser)
    Set pAuthor = value
End Property
