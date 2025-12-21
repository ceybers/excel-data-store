Attribute VB_Name = "RangeHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers.Range"
Option Explicit

'@Description "Returns a new Range with the same shape as the specified 2-dimensional array, starting, from the top-left cell in the specified Range. Throws an error if the input Range is Nothing or if the array is not 2-dimensional."
Public Function ResizeRangeToArray(ByVal InputRange As Range, ByVal InputArray As Variant) As Range
Attribute ResizeRangeToArray.VB_Description = "Returns a new Range with the same shape as the specified 2-dimensional array, starting, from the top-left cell in the specified Range. Throws an error if the input Range is Nothing or if the array is not 2-dimensional."
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputArray)
    Set ResizeRangeToArray = InputRange.Cells.Item(1, 1).Resize(UBound(InputArray, 1), UBound(InputArray, 2))
End Function

'@Description "Updates the Value2 property of the cells in a Range with the values from a 2-dimensional Variant array. If the array is smaller than the Range, only the cells from the top-left to the extents of the array will be updated. If the Range is larger than the Range, the function will update cells outside of the given Range."
Public Sub RangeSetValueFromVariant(ByVal InputRange As Range, ByVal InputVariantArray As Variant)
Attribute RangeSetValueFromVariant.VB_Description = "Updates the Value2 property of the cells in a Range with the values from a 2-dimensional Variant array. If the array is smaller than the Range, only the cells from the top-left to the extents of the array will be updated. If the Range is larger than the Range, the function will update cells outside of the given Range."
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputVariantArray)
    
    Log.Message "RangeSetValueFromVariant writing to = " & InputRange.Address(False, False) & " with Variant(" & UBound(InputVariantArray, 1) & " to " & UBound(InputVariantArray, 2) & ")"
    Dim OutputRange As Range
    Set OutputRange = InputRange.Cells.Item(1, 1).Resize(UBound(InputVariantArray, 1), UBound(InputVariantArray, 2))
    
    Dim VisibleRange As Range
    On Error Resume Next
    Set VisibleRange = OutputRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If VisibleRange Is Nothing Then
        VisibleRange.Value2 = InputVariantArray
        Exit Sub ' All the rows were hidden/filtered-hidden
    ElseIf VisibleRange.Areas.Count = 1 Then
        If VisibleRange.Address = InputRange.Address Then
            VisibleRange.Value2 = InputVariantArray
            Exit Sub ' All the rows were visible
        End If
    End If
    
    Dim VisibleAreaCount As Long
    VisibleAreaCount = VisibleRange.Areas.Count
    
    Dim OptimisticAreaCount As Long
    OptimisticAreaCount = (VisibleRange.Areas.Count * 2) + 1
    
    Dim Extents As Variant
    ReDim Extents(1 To OptimisticAreaCount, 1 To 2)
    
    Extents(1, 1) = OutputRange.Cells.Item(1, 1).Row
    Extents(1, 2) = VisibleRange.Areas.Item(1).Row - 1
    
    Dim Cursor As Long
    Cursor = 1
    
    Dim i As Long
    For i = 1 To VisibleAreaCount
        Dim ThisVisibleArea As Range
        Set ThisVisibleArea = VisibleRange.Areas.Item(i)
        
        If ThisVisibleArea.Row > Extents(Cursor, 1) Then
            Extents(Cursor, 2) = ThisVisibleArea.Row - 1
            
            Cursor = Cursor + 1
            Extents(Cursor, 1) = ThisVisibleArea.Row
            Extents(Cursor, 2) = ThisVisibleArea.Row - 1 + ThisVisibleArea.Rows.Count
            
            Cursor = Cursor + 1
            Extents(Cursor, 1) = Extents(Cursor - 1, 2) + 1
        End If
    Next i
    
    Extents(Cursor, 2) = OutputRange.Row - 1 + OutputRange.Rows.Count
    If Extents(Cursor, 1) > Extents(Cursor, 2) Then Extents(Cursor, 1) = Extents(Cursor, 2)
    
    Dim ActualExtents As Variant
    ReDim ActualExtents(1 To Cursor, 1 To 2)

    Dim Offset As Long
    Offset = Extents(1, 1)
    
    Dim RangeWidth As Long
    RangeWidth = UBound(InputVariantArray, 2)
    
    For i = 1 To Cursor
        Dim Height As Long
        Height = Extents(i, 2) + 1 - Extents(i, 1)
        
        Dim ThisRange As Range
        Set ThisRange = OutputRange.Offset(Extents(i, 1) - Offset, 0).Resize(Height, RangeWidth)
        
        Dim ThisArray As Variant
        ThisArray = ArrayBox(InputVariantArray, Extents(i, 1) + 1 - Offset, 1, Height, RangeWidth)
        
        ThisRange.Value2 = ThisArray
    Next i
End Sub

'@Description "Returns a new Range offset and resized from specified input Range. Returns Nothing if the input Range is nothing. Throws an error if any of the indices are zero or negative."
Public Function RangeBox(ByVal InputRange As Range, ByVal Row As Long, ByVal Column As Long, _
    ByVal Rows As Long, ByVal Columns As Long) As Range
Attribute RangeBox.VB_Description = "Returns a new Range offset and resized from specified input Range. Returns Nothing if the input Range is nothing. Throws an error if any of the indices are zero or negative."
    ' Row = 1 and Column = 1 start the box from the top-left cell.
    ' e.g., RangeBox(Range("A1"), 1, 2, 4, 8).Address = B1:I4
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert Row > 0
    Debug.Assert Column > 0
    Debug.Assert Rows > 0
    Debug.Assert Columns > 0
    
    Set RangeBox = InputRange.Cells.Item(1, 1).Offset(Row - 1, Column - 1).Resize(Rows, Columns)
End Function

' Partitions a Range based on the Values in the given Column.
' CAUTION: Function will Sort rows in the Worksheet.
' Returns a 2-dimensional array with the following:
'  Result(i, 1) is the Value in the Column used to partition
'  Result(i, 2) is the first row this value appears
'  Result(i, 3) is the last row this value appears
'  Result(i, 4) is the Range of the partition
'  Error values in the Partitioning Column will be replaced with Empty.
Public Function PartitionRange(ByVal Range As Range, ByVal Column As Long) As Variant
    If Range Is Nothing Then Exit Function
    If Range.Areas.Count <> 1 Then Exit Function
    If Range.Rows.Count = 1 Then Exit Function
    If Not Range.ListObject Is Nothing Then Exit Function
    
    If Column < 1 Then Exit Function
    If Column > Range.Columns.Count Then Exit Function
    
    Range.Sort Key1:=Range.Columns.Item(Column), Order1:=xlAscending, Header:=xlNo
    
    Dim vv As Variant
    vv = Range.Columns.Item(Column).Value2
    
    Dim i As Long
    For i = 1 To UBound(vv, 1)
        If VarType(vv(i, 1)) = vbError Then
            vv(i, 1) = Empty
        End If
    Next i
    
    Dim Partitions As Variant
    ReDim Partitions(1 To Range.Rows.Count, 1 To 3)
    
    Dim Cursor As Long
    Cursor = 1
    
    Partitions(Cursor, 1) = vv(1, 1)
    Partitions(Cursor, 2) = 1
    
    For i = 2 To Range.Rows.Count
        If vv(i - 1, 1) <> vv(i, 1) Then
            Partitions(Cursor, 3) = i - 1
            Cursor = Cursor + 1
            Partitions(Cursor, 1) = vv(i, 1)
            Partitions(Cursor, 2) = i
        End If
    Next i
    
    Partitions(Cursor, 3) = i - 1
    
    Dim Partitions2 As Variant
    ReDim Partitions2(1 To Cursor, 1 To 4)
    
    For i = 1 To Cursor
        Partitions2(i, 1) = Partitions(i, 1)
        Partitions2(i, 2) = Partitions(i, 2)
        Partitions2(i, 3) = Partitions(i, 3)
        Set Partitions2(i, 4) = Range.Cells.Item(Partitions(i, 2), 1).Resize(Partitions(i, 3) - Partitions(i, 2) + 1, Range.Columns.Count)
    Next i
    
    PartitionRange = Partitions2
End Function

'@Description "Returns True if the Selection object is of type Range and sets the variable to the Range object. Returns False if Selection is Nothing or is not a Range."
Public Function TryGetSelectionRange(ByRef OutRange As Range) As Boolean
Attribute TryGetSelectionRange.VB_Description = "Returns True if the Selection object is of type Range and sets the variable to the Range object. Returns False if Selection is Nothing or is not a Range."
    If Selection Is Nothing Then Exit Function
    If Not TypeOf Selection Is Range Then Exit Function
    
    Set OutRange = Selection
    TryGetSelectionRange = True
End Function

'@Description "Returns True if the two specified Ranges can be intersected and sets the output variable to the intersected Range. Returns False if they cannot be intersected or if one or both of the Range are Nothing."
Public Function TryIntersectRanges(ByVal Range1 As Range, ByVal Range2 As Range, ByRef OutRange As Range) As Boolean
Attribute TryIntersectRanges.VB_Description = "Returns True if the two specified Ranges can be intersected and sets the output variable to the intersected Range. Returns False if they cannot be intersected or if one or both of the Range are Nothing."
    If Range1 Is Nothing Then Exit Function
    If Range2 Is Nothing Then Exit Function
    
    Dim Result As Range
    On Error Resume Next
    Set Result = Application.Intersect(Range1, Range2)
    On Error GoTo 0
    
    If Result Is Nothing Then Exit Function
    
    Set OutRange = Result
    TryIntersectRanges = True
End Function

'@Description "Returns the Value2 property array of a non-contiguous Range that has multiple Areas. The output Variant array is of the same shape as the BaseRange parameter. Cells that are not in the SelectedRange will be Empty variants."
Public Function GetStaggeredArrayValues(ByVal BaseRange As Range, ByVal SelectedRange As Range) As Variant
Attribute GetStaggeredArrayValues.VB_Description = "Returns the Value2 property array of a non-contiguous Range that has multiple Areas. The output Variant array is of the same shape as the BaseRange parameter. Cells that are not in the SelectedRange will be Empty variants."
    If BaseRange Is Nothing Then Exit Function
    If BaseRange.Cells.Count <= 1 Then Exit Function
    If SelectedRange Is Nothing Then Exit Function
    
    On Error Resume Next
    Set SelectedRange = SelectedRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If SelectedRange Is Nothing Then Exit Function
    
    Dim WorkingRange As Range
    If Not TryIntersectRanges(BaseRange, SelectedRange, WorkingRange) Then Exit Function
    
    Dim WorksheetOffset As Long
    WorksheetOffset = BaseRange.Cells.Item(1, 1).Row - 1
    
    Dim BaseRangeValues As Variant
    BaseRangeValues = BaseRange.Value2
    
    Dim Result As Variant
    ReDim Result(1 To UBound(BaseRangeValues, 1), 1 To 1)

    Dim AreaIndex As Long
    For AreaIndex = 1 To WorkingRange.Areas.Count
        Dim Area As Range
        Set Area = WorkingRange.Areas.Item(AreaIndex)
        Dim RowIndex As Long
        For RowIndex = 1 To Area.Rows.Count
            Dim ThisRowIndex As Long
            ThisRowIndex = Area.Rows.Item(RowIndex).Row - WorksheetOffset
            Result(ThisRowIndex, 1) = BaseRangeValues(ThisRowIndex, 1)
        Next RowIndex
    Next AreaIndex
    
    ArrayClean.ReplaceErrorCells2 Result
    
    GetStaggeredArrayValues = Result
End Function
