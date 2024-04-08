VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBreakDownTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pStartingPosition As Range
Private pCellvalue As String
Private pCellformat As String

Public Property Get startingPosition() As Range
    Set startingPosition = pStartingPosition
End Property

Public Property Let startingPosition(value As Range)
    Set pStartingPosition = value
End Property

Public Property Get cellvalue() As String
    cellvalue = pCellvalue
End Property

Public Property Let cellvalue(value As String)
    pCellvalue = value
End Property

Public Property Get cellformat() As String
    cellformat = pCellformat
End Property

Public Property Let cellformat(value As String)
    pCellformat = value
End Property

Private Sub Class_initialize()
    startingPosition = Range(ActiveCell.Address)
End Sub

Public Function Cell(Optional row As Integer = 0, Optional col As Integer = 0)
    Set Cell = Cells(pStartingPosition.row + row, pStartingPosition.Column + col)
End Function

Public Function GetCellPosition(Optional row As Integer = 0, Optional col As Integer = 0) As Range
    Set GetCellPosition = Range(Cell(row, col).Address)
End Function


