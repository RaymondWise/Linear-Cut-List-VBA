Attribute VB_Name = "CutList"
Option Explicit
Global Const BOARD_LENGTH As Long = 95
Public Sub DimensionalLumberCutList()

    Dim lastRow As Long
    lastRow = Sheet1.Cells(Rows.Count, 1).End(xlUp).Row
    Dim listofcomponents() As Double
    listofcomponents = GetListOfComponents(lastRow)
    
    Dim totalLength As Long
    totalLength = GetTotalLength(listofcomponents)
    Dim minimumNumberOfBoards As Long
    minimumNumberOfBoards = Application.WorksheetFunction.RoundUp(totalLength / BOARD_LENGTH, 0)
    
    Dim lumberStack As Collection
    Set lumberStack = New Collection
    Dim board As TwoByFour
    Dim boardCount As Long
    Dim listIndex As Long
    For boardCount = 1 To minimumNumberOfBoards
        Set board = New TwoByFour
            For listIndex = LBound(listofcomponents) To UBound(listofcomponents)
                If board.Offcut < listofcomponents(UBound(listofcomponents)) Then
                    lumberStack.Add board
                    Exit For
                End If
                If board.Offcut > listofcomponents(listIndex) And listofcomponents(listIndex) <> 0 Then
                    board.MakeCut listofcomponents(listIndex)
                    listofcomponents(listIndex) = 0
                End If
                If listIndex = UBound(listofcomponents) Then
                    lumberStack.Add board
                    Exit For
                End If
            Next
            'some sort of error here***************************************
            If Not Application.WorksheetFunction.Sum(listofcomponents) = 0 And boardCount = minimumNumberOfBoards Then
                minimumNumberOfBoards = minimumNumberOfBoards + 1
            End If
    Next
    With Sheet2
        .UsedRange.Clear
        For boardCount = 1 To lumberStack.Count
            .Range(.Cells(boardCount, 1), .Cells(boardCount, lumberStack(boardCount).NumberOfCuts)) = lumberStack(boardCount).WriteCuts
        Next
    End With

End Sub


Private Function GetListOfComponents(ByVal lastRow As Long) As Double()
    Dim componentDataArray As Variant
    componentDataArray = PopulateComponentDataArray(lastRow)
    Dim numberOfComponents As Long
    numberOfComponents = GetNumberOfComponents(componentDataArray)
    Dim componentDoubleArray() As Double
    ReDim componentDoubleArray(1 To numberOfComponents)
    Dim index As Long
    index = 1
    Dim counter As Long
    Dim quantityOfEach As Long
    For counter = 1 To lastRow - 1
        For quantityOfEach = 1 To componentDataArray(counter, 2)
            componentDoubleArray(index) = componentDataArray(counter, 1)
            index = index + 1
        Next
    Next
    CombSortNumbers componentDoubleArray
    GetListOfComponents = componentDoubleArray
End Function

Private Function PopulateComponentDataArray(ByVal lastRow As Long) As Variant
    Dim componentRange As Range
    Set componentRange = Sheet1.Range(Sheet1.Cells(2, 1), Sheet1.Cells(lastRow, 2))
    PopulateComponentDataArray = componentRange
End Function

Private Function GetNumberOfComponents(ByVal componentDataArray As Variant) As Long
    Dim counter As Long
    For counter = LBound(componentDataArray) To UBound(componentDataArray)
        GetNumberOfComponents = GetNumberOfComponents + componentDataArray(counter, 2)
    Next
End Function

Private Function GetTotalLength(ByVal listofcomponents As Variant) As Double
    Dim index As Long
    For index = LBound(listofcomponents) To UBound(listofcomponents)
        GetTotalLength = GetTotalLength + listofcomponents(index)
    Next
End Function

Private Sub CombSortNumbers(ByRef numberArray() As Double, Optional ByVal sortAscending As Boolean = False)
    Const SHRINK As Double = 1.3
    Dim initialSize As Long
    initialSize = UBound(numberArray())
    Dim gap As Long
    gap = initialSize
    Dim index As Long
    Dim isSorted As Boolean

    Do While gap > 1 And Not isSorted
        gap = Int(gap / SHRINK)
        If gap > 1 Then
            isSorted = False
        Else
            gap = 1
            isSorted = True
        End If
        index = LBound(numberArray)
        Do While index + gap <= initialSize
            If sortAscending Then
                If numberArray(index) > numberArray(index + gap) Then
                    SwapElements numberArray, index, index + gap
                    isSorted = False
                End If
            Else
                If numberArray(index) < numberArray(index + gap) Then
                    SwapElements numberArray, index, index + gap
                    isSorted = False
                End If
            End If
            index = index + 1
        Loop
    Loop

End Sub

Private Sub SwapElements(ByRef numberArray() As Double, ByVal i As Long, ByVal j As Long)
    Dim temporaryHolder As Double
    temporaryHolder = numberArray(i)
    numberArray(i) = numberArray(j)
    numberArray(j) = temporaryHolder
End Sub
