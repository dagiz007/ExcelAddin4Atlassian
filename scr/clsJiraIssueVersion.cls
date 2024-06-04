VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJiraIssueVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pId As String
Private pReleaseDate As Date
Private pVersionName As String
Private pArchived As Boolean
Private pReleased As Boolean

Public Property Get Id() As String
    Id = pId
End Property

Public Property Let Id(value As String)
    pId = value
End Property

Public Property Get releaseDate() As Date
    releaseDate = pReleaseDate
End Property

Public Property Let releaseDate(value As Date)
    pReleaseDate = value
End Property

Public Property Get versionName() As String
    versionName = pVersionName
End Property

Public Property Let versionName(value As String)
    pVersionName = value
End Property

Public Property Get archived() As Boolean
    archived = pArchived
End Property

Public Property Let archived(value As Boolean)
    pArchived = value
End Property

Public Property Get released() As Boolean
    released = pReleased
End Property

Public Property Let released(value As Boolean)
    pReleased = value
End Property
