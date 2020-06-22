VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "comment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pCreatedDate As Date
Private pBody As String
Private pAuthor As user

Public Property Get createdDate() As Date
    createdDate = pCreatedDate
End Property

Public Property Let createdDate(value As Date)
    pCreatedDate = value
End Property

Public Property Get body() As String
    body = pBody
End Property

Public Property Let body(value As String)
    pBody = value
End Property

Public Property Get author() As user
    Set author = pAuthor
End Property

Public Property Let author(value As user)
    Set pAuthor = value
End Property


