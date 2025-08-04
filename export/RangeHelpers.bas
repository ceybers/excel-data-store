Attribute VB_Name = "RangeHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers.Range"
Option Explicit

' Returns a range with the same shape as the specified 2-dimensional InputArray, starting
' from the top-most cell in the specified InputRange.
Public Function ResizeRangeToArray(ByVal InputRange As Range, ByVal InputArray As Variant) As Range
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputArray)
    Set ResizeRangeToArray = InputRange.Cells.Item(1, 1).Resize(UBound(InputArray, 1), UBound(InputArray, 2))
End Function

' Updates the .Value2 property of all the cells in the InputRange with the Variant Values
' in the specified 2-dimensional InputArray.
Public Sub RangeSetValueFromVariant(ByVal InputRange As Range, ByVal InputVariant As Variant)
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputVariant)
    InputRange.Cells.Item(1, 1).Resize(UBound(InputVariant, 1), UBound(InputVariant, 2)).Value2 = InputVariant
End Sub

' Returns a range with the offset and size of the specified input parameters, starting from the
' top-most cell in the InputRange. Row = 1 and Column = 1 start the box from the top-left cell.
' e.g., RangeBox(Range("A1"), 1, 2, 4, 8).Address = B1:I4
Public Function RangeBox(ByVal InputRange As Range, ByVal Row As Long, ByVal Column As Long, _
    ByVal Rows As Long, ByVal Columns As Long) As Range
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert Row > 0
    Debug.Assert Column > 0
    Debug.Assert Rows > 0
    Debug.Assert Columns > 0
    
    Set RangeBox = InputRange.Cells.Item(1, 1).Offset(Row - 1, Column - 1).Resize(Rows, Columns)
End Function

