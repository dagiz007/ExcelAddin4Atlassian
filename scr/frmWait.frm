VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWait 
   Caption         =   "ExcelAddin4Atlassian"
   ClientHeight    =   1740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4020
   OleObjectBlob   =   "frmWait.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbort_Click()
    If vbYes = MsgBox("Are you sure you want to abort.", vbInformation + vbYesNo) Then End
End Sub

Private Sub lblLink_Click()
 Call OpenHyperlink("https://github.com/dagiz007/ExcelAddin4Atlassian")
End Sub
