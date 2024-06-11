VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAtlassianUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pAccountId As String
Private pEmailAddress As String
Private pDisplayName As String
Private pActive As Boolean
Private pLastActive As Date
Private pProductAccess As Collection

Public Property Get accountId() As String
    accountId = pAccountId
End Property

Public Property Let accountId(value As String)
    pAccountId = value
End Property

Public Property Get email() As String
    email = pEmailAddress
End Property

Public Property Let email(value As String)
    pEmailAddress = value
End Property

Public Property Get name() As String
    name = pDisplayName
End Property

Public Property Let name(value As String)
    pDisplayName = value
End Property

Public Property Get active() As Boolean
    active = pActive
End Property

Public Property Let active(value As Boolean)
    pActive = value
End Property

Public Property Get lastActive() As Date
    lastActive = pLastActive
End Property

Public Property Let lastActive(value As Date)
    pLastActive = value
End Property

Public Property Get productAccess() As Collection
    Set productAccess = pProductAccess
End Property

Public Property Let productAccess(value As Collection)
    Set pProductAccess = value
End Property
