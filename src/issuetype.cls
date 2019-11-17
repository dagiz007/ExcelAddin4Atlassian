VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "issuetype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private pName As String
Private pSubtask As Boolean

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(value As String)
    pId = value
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let name(value As String)
    pName = value
End Property

Public Property Get subtask() As Boolean
    subtask = sSubtask
End Property

Public Property Let subtask(value As Boolean)
    pSubtask = value
End Property
