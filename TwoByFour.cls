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
    Const DEFAULT_LENGTH As Double = 96
    Const DEFAULT_KERF As Double = 1 / 16
    
    Private remainder As Double
    Private listOfCuts As Object
    Private cutKerf As Double
    Private startLength As Double
    
    Public Property Get numberOfCuts() As Long
        numberOfCuts = listOfCuts.Count
    End Property
    Public Property Get Offcut() As Double
        Offcut = remainder
    End Property
    Public Property Get bladeKerf() As Double
        bladeKerf = cutKerf
    End Property
    Public Property Get boardLength() As Double
        boardLength = startLength
    End Property
    
    Public Property Let boardLength(ByVal inputLength As Double)
        If listOfCuts.Count > 1 Then Exit Property
        startLength = inputLength
    End Property
    Public Property Let bladeKerf(ByVal inputKerf As Double)
        cutKerf = inputKerf
    End Property
    
    Private Sub Class_Initialize()
        Set listOfCuts = CreateObject("Scripting.Dictionary")
        remainder = DEFAULT_LENGTH
        cutKerf = DEFAULT_KERF
        startLength = remainder
    End Sub
    
    Public Sub MakeCut(ByVal component As String, ByVal length As Double)
        remainder = remainder - length - cutKerf
        listOfCuts.Add component, length
    End Sub
    Public Property Get CutArray() As Variant
        CutArray = WriteCuts
    End Property
    Public Function WriteCuts() As Variant
        Dim cutCount As Long
        cutCount = listOfCuts.Count
        Dim tempArray As Variant
        ReDim tempArray(1 To cutCount)
        Dim i As Long
        i = 1
        Dim key As Variant
        For Each key In listOfCuts.keys
                tempArray(i) = "Component " & key & " @ " & listOfCuts.Item(key) & "''"
                i = i + 1
        Next
        WriteCuts = tempArray
    End Function
