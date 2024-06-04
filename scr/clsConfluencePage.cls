VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfluencePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pId As String
Private pTitle As String
Private pVersion As Integer
Private pBody As String
Private pSpaceId As String
Private pStatus As String

Public Property Get Id() As String
    Id = pId
End Property

Public Property Let Id(value As String)
    pId = value
End Property

Public Property Get title() As String
    title = pTitle
End Property

Public Property Let title(value As String)
    pTitle = value
End Property

Public Property Get version() As Integer
    version = pVersion
End Property

Public Property Let version(value As Integer)
    pVersion = value
End Property

Public Property Get body() As String
    body = pBody
End Property

Public Property Let body(value As String)
    pBody = value
End Property

Public Property Get spaceId() As String
    spaceId = pSpaceId
End Property

Public Property Let spaceId(value As String)
    pSpaceId = value
End Property

Public Property Get status() As String
    status = pStatus
End Property

Public Property Let status(value As String)
    pStatus = value
End Property

