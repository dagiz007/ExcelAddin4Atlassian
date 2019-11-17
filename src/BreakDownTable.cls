VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BreakDownTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pStartingPosition As Range

Public Property Get StartingPosition() As Range
    Set StartingPosition = pStartingPosition
End Property

Public Property Let StartingPosition(value As Range)
    Set pStartingPosition = value
End Property

Public Function Cell(row As Integer, col As Integer)
    Set Cell = Cells(pStartingPosition.row + row, pStartingPosition.Column + col)
End Function
