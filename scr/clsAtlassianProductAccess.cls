VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAtlassianProductAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pName As String
Private pUrl As String
Private pLastActive As Date

Public Property Get name() As String
    name = pName
End Property

Public Property Let name(value As String)
    pName = value
End Property

Public Property Get url() As String
    url = pUrl
End Property

Public Property Let url(value As String)
    pUrl = value
End Property

Public Property Get lastActive() As Date
    lastActive = pLastActive
End Property

Public Property Let lastActive(value As Date)
    pLastActive = value
End Property
