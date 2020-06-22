VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "worklog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private pTimeSpentSeconds As String
Private pStarted As Date
Private pAuthor As user

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(value As String)
    pId = value
End Property

Public Property Get timeSpentSeconds() As String
    timeSpentSeconds = pTimeSpentSeconds
End Property

Public Property Let timeSpentSeconds(value As String)
    pTimeSpentSeconds = value
End Property

Public Property Get started() As Date
    started = pStarted
End Property

Public Property Let started(value As Date)
    pStarted = value
End Property

Public Property Get author() As user
    Set author = pAuthor
End Property

Public Property Let author(value As user)
    Set pAuthor = value
End Property

