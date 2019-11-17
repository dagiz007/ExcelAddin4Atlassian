VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "transition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pTransitionDate As Date
Private pFromId As String
Private pFromString As String
Private pToId As String
Private pToString As String
Private pTimeInSourceStatus As Integer

Public Property Get transitionDate() As Date
    transitionDate = pTransitionDate
End Property

Public Property Let transitionDate(value As Date)
    pTransitionDate = value
End Property

Public Property Get fromId() As String
    fromId = pFromId
End Property

Public Property Let fromId(value As String)
    pFromId = value
End Property

Public Property Get fromString() As String
    fromString = pFromString
End Property

Public Property Let fromString(value As String)
    pFromString = value
End Property

Public Property Get toId() As String
    toId = pToId
End Property

Public Property Let toId(value As String)
    pToId = value
End Property

Public Property Get toString() As String
    toString = pToString
End Property

Public Property Let toString(value As String)
    pToString = value
End Property

Public Property Get timeInSourceStatus() As Integer
    timeInSourceStatus = pTimeInSourceStatus
End Property

Public Property Let timeInSourceStatus(value As Integer)
    pTimeInSourceStatus = value
End Property
