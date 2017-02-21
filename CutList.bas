Attribute VB_Name = "CutList"
Option Explicit
Public Sub DimensionalLumberCutList()
    Dim sourceSheet As Worksheet
    Set sourceSheet = Sheet1
    Dim targetSheet As Worksheet
    Set targetSheet = Sheet2
    
    Dim lastRow As Long
    lastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim listOfComponents As Variant
    listOfComponents = GetListOfComponents(sourceSheet, lastRow)
    
    Dim lumberStack As Collection
    Set lumberStack = GetCutlist(listOfComponents, sourceSheet)
    
    PrintCuts targetSheet, lumberStack
    
End Sub

Private Function GetListOfComponents(ByVal sourceSheet As Worksheet, ByVal lastRow As Long) As Variant
    Dim inputData As Variant
    inputData = PopulateinputData(sourceSheet, lastRow)
    Dim numberOfComponents As Long
    numberOfComponents = GetNumberOfComponents(inputData)
    Dim componentArray As Variant
    ReDim componentArray(1 To numberOfComponents, 1 To 2)
    Dim index As Long
    index = 1
    Dim counter As Long
    Dim quantityOfEach As Long
    For counter = 1 To lastRow - 1
        For quantityOfEach = 1 To inputData(counter, 2)
            componentArray(index, 1) = inputData(counter, 1)
            componentArray(index, 2) = inputData(counter, 1) & "-" & quantityOfEach
            index = index + 1
        Next
    Next
    CombSortArray componentArray, 2, 1
    GetListOfComponents = componentArray
End Function

Private Function PopulateinputData(ByVal sourceSheet As Worksheet, ByVal lastRow As Long) As Variant
    Dim componentRange As Range
    Set componentRange = sourceSheet.Range(sourceSheet.Cells(2, 1), sourceSheet.Cells(lastRow, 2))
    PopulateinputData = componentRange
End Function

Private Function GetNumberOfComponents(ByVal inputData As Variant) As Long
    Dim counter As Long
    For counter = LBound(inputData) To UBound(inputData)
        GetNumberOfComponents = GetNumberOfComponents + inputData(counter, 2)
    Next
End Function

Private Function GetCutlist(ByRef listOfComponents As Variant, ByVal sourceSheet As Worksheet) As Collection
    Dim lumberStack As Collection
    Set lumberStack = New Collection
    Dim board As TwoByFour
    Dim index As Long
    
    Do
        Set board = New TwoByFour
        board.boardLength = sourceSheet.Range("boardLength")
        board.bladeKerf = sourceSheet.Range("bladeKerf")
        For index = LBound(listOfComponents, 1) To UBound(listOfComponents, 1)
            If board.Offcut < listOfComponents(UBound(listOfComponents, 1), 1) Then Exit For
            If board.Offcut > listOfComponents(index, 1) And listOfComponents(index, 1) <> 0 Then
                board.MakeCut CStr(listOfComponents(index, 2)), CDbl(listOfComponents(index, 1))
                listOfComponents(index, 1) = 0
            End If
        Next
        lumberStack.Add board
    Loop While Application.WorksheetFunction.Sum(Application.WorksheetFunction.index(listOfComponents, 0, 1)) > 0
    Set GetCutlist = lumberStack
End Function

Private Sub PrintCuts(ByVal targetSheet As Worksheet, ByVal lumberStack As Collection)
    Const PROJECT_REQUIRES As String = "Your project requires "
    Const BOARDS_REQUIRED As String = "''" & " boards:"
    Const OFFCUT_LENGTH As String = "Offcut: "
    Const BOARD_TITLE As String = "Board "

    Dim arrayOfCuts As Variant
    Dim lastColumn As Long
    Dim index As Long
    With targetSheet
        .UsedRange.Clear
        .Cells(1, 1) = PROJECT_REQUIRES & lumberStack.Count & " " & lumberStack(1).boardLength & BOARDS_REQUIRED
        For index = 1 To lumberStack.Count
            arrayOfCuts = lumberStack(index).CutArray
            .Cells(index + 1, 1) = BOARD_TITLE & index
            .Range(.Cells(index + 1, 2), .Cells(index + 1, UBound(arrayOfCuts) + 1)) = arrayOfCuts
            .Cells(index + 1, UBound(arrayOfCuts) + 3) = OFFCUT_LENGTH & lumberStack(index).Offcut
            If lastColumn < UBound(arrayOfCuts) + 3 Then lastColumn = UBound(arrayOfCuts) + 3
        Next
    End With
    FormatList targetSheet, lumberStack.Count + 1, lastColumn, lumberStack(1).bladeKerf
End Sub

Private Sub FormatList(ByVal targetSheet As Worksheet, ByVal lastRow As Long, ByVal lastColumn As Long, ByVal bladeKerf As Double)
    Const KERF_WASTE As String = "Kerf waste: "
    Dim targetRow As Long
    Dim offCutColumn As Long
    With targetSheet
        For targetRow = 2 To lastRow
            offCutColumn = .Cells(targetRow, Columns.Count).End(xlToLeft).Column
            If Not offCutColumn = lastColumn Then
                .Cells(targetRow, lastColumn) = .Cells(targetRow, offCutColumn)
                .Cells(targetRow, offCutColumn) = vbNullString
            End If
            .Cells(targetRow, lastColumn + 1) = KERF_WASTE & (offCutColumn - 3) * bladeKerf
        Next
        
        .Columns.AutoFit
        .Columns(1).ColumnWidth = 10
    End With
End Sub


Private Sub CombSortArray(ByRef dataArray As Variant, Optional ByVal numberOfColumns As Long = 1, Optional ByVal sortKeyColumn As Long = 1)
    'Comb Sort procedure only sorts descending because cutlist is relying on descending values
    'For full procedure see https://github.com/RaymondWise/VBACombSort
    Const SHRINK As Double = 1.3
    
    Dim initialSize As Long
    initialSize = UBound(dataArray, 1)
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
        index = 1
        Do While index + gap <= initialSize
            If dataArray(index, sortKeyColumn) < dataArray(index + gap, sortKeyColumn) Then
               SwapElements dataArray, numberOfColumns, index, index + gap
               isSorted = False
            End If
            index = index + 1
        Loop
    Loop

End Sub

Private Sub SwapElements(ByRef dataArray As Variant, ByVal numberOfColumns As Long, ByVal i As Long, ByVal j As Long)
    Dim temporaryHolder As Variant
    Dim index As Long
    For index = 1 To numberOfColumns
        temporaryHolder = dataArray(i, index)
        dataArray(i, index) = dataArray(j, index)
        dataArray(j, index) = temporaryHolder
    Next
End Sub
