VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "attachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private pFileName As String

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(value As String)
    pId = value
End Property

Public Property Get fileName() As String
    fileName = pFileName
End Property

Public Property Let fileName(value As String)
    pFileName = value
End Property
