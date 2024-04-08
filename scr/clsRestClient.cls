VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRestClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pStatus As String
Private pResponseText As String
Private pResponseBody As Variant

Public Property Get status() As String
    status = pStatus
End Property

Public Property Let status(value As String)
    pStatus = value
End Property

Public Property Get responseText() As String
    responseText = pResponseText
End Property

Public Property Let responseText(value As String)
    pResponseText = value
End Property

Public Property Get responseBody() As Variant
    responseBody = pResponseBody
End Property

Public Property Let responseBody(value As Variant)
    pResponseBody = value
End Property

Private Sub Class_initialize()
    If gstrBoundary = vbNullString Then gstrBoundary = CreateBoundary
End Sub

Public Function Service(method As String, url As String, Optional payload As Variant = "") As clsRestClient

    On Error GoTo ErrHandler
           
    Dim RestClient As Object
    Set RestClient = CreateObject("WinHttp.WinHttpRequest.5.1")

     With RestClient
        .Open method, gstrAtlassianURL & url, False
        .setRequestHeader "Authorization", "Basic " & EncodeBase64(gstrAtlassianEmail & ":" & gstrAtlassianToken)
    
        If VarType(payload) = 8209 Then
            .setRequestHeader "X-Atlassian-Token", "nocheck"
            .setRequestHeader "Content-Type", "multipart/form-data; charset=""ISO-8859-1""; boundary=" & gstrBoundary
        Else
            .setRequestHeader "Accept", "application/json; charset=utf-8"
            .setRequestHeader "Content-Type", "application/json"
        End If
        
        .Send payload
        
        Set Service = New clsRestClient
        Service.status = .status
        Service.responseText = .responseText
        Service.responseBody = .responseBody
        
        If Not Service.status = 200 Then
            Debug.Print Service.responseText
        Else
            'Debug.Print Service.responseText
        End If
        
        If gblnLogging Then Call addLog(Format(Now(), "YYYYMMDD") & "_RestAudit.log", gstrAtlassianURL & url & vbTab & _
                                         method & vbTab & _
                                         payload & vbTab & _
                                         Service.status & vbTab & _
                                         Service.responseText)
    End With
        
    Set RestClient = Nothing
    
    Set Service = Service
    
    Exit Function


ErrHandler:

    MsgBox Err.description, vbCritical, Err.Number
    End
    
End Function

Private Function CreateBoundary() As String

    Dim characters As String
    Dim boundary As String
    Dim i As Integer
    
    characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    boundary = ""
        
    For i = 1 To 20
        boundary = boundary & Mid(characters, Int(Len(characters) * Rnd + 1), 1)
    Next
    
    CreateBoundary = boundary

End Function

Private Function EncodeBase64(text As String) As String

    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
  
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
  
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
  
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = Replace(objNode.text, vbLf, "")
  
    Set objNode = Nothing
    Set objXML = Nothing
    
End Function
