VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pUserKey As String
Private pEmailAddress As String
Private pDisplayName As String
Private pActive As Boolean

Public Property Get userKey() As String
    userKey = pUserKey
End Property

Public Property Let userKey(value As String)
    pUserKey = value
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
