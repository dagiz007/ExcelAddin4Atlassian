VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfluenceSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pId As String
Private pKey As String
Private pSpaceType As String
Private pName As String

Public Property Get Id() As String
    Id = pId
End Property

Public Property Let Id(value As String)
    pId = value
End Property

Public Property Get key() As String
    key = pKey
End Property

Public Property Let key(value As String)
    pKey = value
End Property

Public Property Get spaceType() As String
    spaceType = pSpaceType
End Property

Public Property Let spaceType(value As String)
    pSpaceType = value
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let name(value As String)
    pName = value
End Property

