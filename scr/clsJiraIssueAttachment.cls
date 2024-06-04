VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJiraIssueAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pId As String
Private pFilename As String
Private pData As Variant

Public Property Get Id() As String
    Id = pId
End Property

Public Property Let Id(value As String)
    pId = value
End Property

Public Property Get filename() As String
    filename = pFilename
End Property

Public Property Let filename(value As String)
    pFilename = value
End Property

Public Property Get data() As Variant
    data = pData
End Property

Public Property Let data(value As Variant)
    pData = value
End Property
