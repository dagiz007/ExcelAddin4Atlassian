VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJiraProjectRoleActor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pRoleType As String
Private pName As String
Private pId As String

Public Property Get roleType() As String
    roleType = pRoleType
End Property

Public Property Let roleType(value As String)
    pRoleType = value
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let name(value As String)
    pName = value
End Property

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(value As String)
    pId = value
End Property
