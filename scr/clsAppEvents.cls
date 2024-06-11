VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents mxlApp As Application       'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
Attribute mxlApp.VB_VarHelpID = -1

Public Property Set App(xlApp As Application)  'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
    Set mxlApp = xlApp
End Property

Private Sub mxlApp_SheetCalculate(ByVal Sh As Object)
    
    If gcolAppEventResult.Count > 0 Then
    
        Dim bdTable As clsBreakDownTable
        Dim selectedPosition As Range
        
        For Each bdTable In gcolAppEventResult
            DoEvents
            
            If selectedPosition Is Nothing Then Set selectedPosition = bdTable.startingPosition

            bdTable.Cell.NumberFormat = bdTable.cellformat
            bdTable.Cell.value = bdTable.cellvalue
        Next
        
        selectedPosition.Select
        Set gcolAppEventResult = New Collection
    
    End If
    
    Unload frmWait

End Sub
