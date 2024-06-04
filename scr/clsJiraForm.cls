VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJiraForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private pInternal As Boolean
Private pSubmitted As Boolean
Private pLoced As Boolean
Private pName As String
Private pUpdated As Date

Public Property Get Id() As String
    Id = pId
End Property

Public Property Let Id(value As String)
    pId = value
End Property

Public Property Get internal() As Boolean
    internal = pInternal
End Property

Public Property Let internal(value As Boolean)
    pInternal = value
End Property

Public Property Get submitted() As Boolean
    submitted = pSubmitted
End Property

Public Property Let submitted(value As Boolean)
    pSubmitted = value
End Property

Public Property Get locked() As Boolean
    locked = pLocked
End Property

Public Property Let locked(value As Boolean)
    pLocked = value
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let name(value As String)
    pName = value
End Property

Public Property Get updated() As Date
    updated = pUpdated
End Property

Public Property Let updated(value As Date)
    pUpdated = value
End Property
