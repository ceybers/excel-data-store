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
