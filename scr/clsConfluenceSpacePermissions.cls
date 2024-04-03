VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfluenceSpacePermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pPermissionType As String
Private pName As String
Private pId As String
Private pEmail As String
Private pOperation As String
Private pTargetType As String

Public Property Get permissionType() As String
    permissionType = pPermissionType
End Property

Public Property Let permissionType(value As String)
    pPermissionType = value
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

Public Property Get email() As String
    email = pEmail
End Property

Public Property Let email(value As String)
    pEmail = value
End Property

Public Property Get operation() As String
    operation = pOperation
End Property

Public Property Let operation(value As String)
    pOperation = value
End Property

Public Property Get targetType() As String
    targetType = pTargetType
End Property

Public Property Let targetType(value As String)
    pTargetType = value
End Property


