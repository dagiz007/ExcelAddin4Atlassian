VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7290
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkLogging_Change()
    If chkLogging Then
        txtLogPath.Enabled = True
    Else
        txtLogPath.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
            
    gstrAtlassianURL = Trim(txtAtlassianURL)
    If Right(gstrAtlassianURL, 1) = "/" Then gstrAtlassianURL = Left(gstrAtlassianURL, Len(gstrAtlassianURL) - 1)
    
    gstrAtlassianEmail = Trim(txtAtlassianEmail)
    gstrAtlassianToken = Trim(txtAtlassianToken)
    
    gstrLogPath = Trim(txtLogPath)
    If Right(gstrLogPath, 1) = "\" Then gstrLogPath = Left(gstrLogPath, Len(gstrLogPath) - 1)
    
    gblnLogging = chkLogging
        

    lblAtlassianURL.ForeColor = vbBlack
    lblAtlassianEmail.ForeColor = vbBlack
    lblAtlassianToken.ForeColor = vbBlack
    lblLogPath.ForeColor = vbBlack
       
    If gstrAtlassianEmail = vbNullString Then
        MultiPage1.value = 0
        lblAtlassianEmail.ForeColor = vbRed
        txtAtlassianEmail.SetFocus
        Exit Sub
    End If

    If gstrAtlassianToken = vbNullString Then
        MultiPage1.value = 0
        lblAtlassianToken.ForeColor = vbRed
        txtAtlassianToken.SetFocus
        Exit Sub
    End If
        
    If gstrAtlassianURL = vbNullString Then
        MultiPage1.value = 0
        lblAtlassianURL.ForeColor = vbRed
        txtAtlassianURL.SetFocus
        Exit Sub
    End If
    
    If chkLogging And (gstrLogPath = vbNullString Or Dir(gstrLogPath, vbDirectory) = "" Or Dir(gstrLogPath, vbDirectory) = ".") Then
        lblLogPath.ForeColor = vbRed
        txtLogPath.SetFocus
        If Not gstrLogPath = vbNullString Then MsgBox "Folder " & gstrLogPath & " does not exist.", vbCritical
        Exit Sub
    End If
    
    
    'Set gblnSuccessfulLogin to True to avoid Jira class_initialize from making the same check and reloading this form.
    gblnSuccessfulLogin = True
    
    If Jira.CorrectCredentianls Then
        SaveSetting "ExcelAddin4Atlassian", "Settings", "AtlassianURL", gstrAtlassianURL
        SaveSetting "ExcelAddin4Atlassian", "Settings", "AtlassianEmail", gstrAtlassianEmail
        SaveSetting "ExcelAddin4Atlassian", "Settings", "AtlassianToken", gstrAtlassianToken
        SaveSetting "ExcelAddin4Atlassian", "Settings", "Logging", gblnLogging
        SaveSetting "ExcelAddin4Atlassian", "Settings", "LogPath", gstrLogPath

        Unload Me
    Else
        gblnSuccessfulLogin = False
        MsgBox "Wrong credentials or URL", vbCritical
    End If

End Sub

Private Sub lblLink_Click()
    Call OpenHyperlink("https://github.com/dagiz007/ExcelAddin4Atlassian")
End Sub

Private Sub UserForm_Initialize()
    
    txtAtlassianURL = gstrAtlassianURL
    txtAtlassianEmail = gstrAtlassianEmail
    txtAtlassianToken = gstrAtlassianToken
    chkLogging = gblnLogging
    txtLogPath = gstrLogPath

    MultiPage1.value = 0
    
End Sub
