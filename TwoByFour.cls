VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TwoByFour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private index As Long
Private remainder As Double
Private listOfCuts() As Double

Private Sub Class_Initialize()
    ReDim listOfCuts(1 To 1)
    listOfCuts(1) = BOARD_LENGTH
End Sub

Public Property Get NumberOfCuts() As Long
    NumberOfCuts = UBound(listOfCuts)
End Property

Public Property Get Offcut() As Double
    Offcut = listOfCuts(UBound(listOfCuts))
End Property

Public Sub MakeCut(length As Double)
    index = UBound(listOfCuts)
    remainder = listOfCuts(index) - length
    listOfCuts(index) = length
    ReDim Preserve listOfCuts(1 To index + 1)
    listOfCuts(index + 1) = remainder
End Sub

Public Function WriteCuts() As Variant
    'WriteCuts = Application.Transpose(listOfCuts)
    WriteCuts = listOfCuts
End Function
