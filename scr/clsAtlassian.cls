VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAtlassian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private RestClient As New clsRestClient

Private Sub Class_initialize()
    If gstrAtlassianOrgId = vbNullString Then Call LoadSettings
    
    If gstrAtlassianOrgId = vbNullString Or gstrAtlassianOrgToken = vbNullString Then
        frmSettings.Show
    End If
End Sub

Private Sub Class_Terminate()
    Set RestClient = Nothing
End Sub

Public Function GetUsers() As Collection

    Dim url As String
    Dim json As Object
    
    Dim item As Variant
    Dim user As clsAtlassianUser
    Dim users As New Collection
    
    Dim productItem As Variant
    Dim product As clsAtlassianProductAccess
    Dim products As Collection
    
    url = "https://api.atlassian.com/admin/v1/orgs/" & gstrAtlassianOrgId & "/users"
    
    Do
        DoEvents
        
        Set RestClient = RestClient.Service("GET", url)

        Set json = JsonConverter.ParseJson(RestClient.responseText)
        
        For Each item In json("data")
            Set user = New clsAtlassianUser
            user.accountId = item("account_id")
            user.name = item("name")
            user.email = item("email")
            If item("account_status") = "active" Then user.active = True
            If Not IsEmpty(item("last_active")) Then user.lastActive = ParseISO8601date(item("last_active"))
            
            Set products = New Collection
            
            For Each productItem In item("product_access")
                Set product = New clsAtlassianProductAccess
                product.name = productItem("name")
                product.url = productItem("url")
                If Not IsEmpty(productItem("last_active")) Then product.lastActive = ParseISO8601date(productItem("last_active"))
                products.Add product
            Next
            
            user.productAccess = products
            users.Add user
        Next
      
        If IsEmpty(json("links")("next")) Then
            Exit Do
        Else
            url = json("links")("next")
        End If
        
    Loop
    
    Set GetUsers = users
    
End Function




