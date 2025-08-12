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
    If Column < 1 Then Exit Function
    If Column > Range.Columns.Count Then Exit Function
    
    ' handle this odd case
    If Range.Rows.Count = 1 Then Exit Function
    
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
    
    Partitions(1, 1) = vv(1, 1)
    Partitions(1, 2) = 1
    
    Dim Cursor As Long
    Cursor = 1
    
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
        'Set Partitions2(i, 4) = Range.Range(Range.Rows(Partitions(i, 2)), Range.Rows(Partitions(i, 3)))
        Set Partitions2(i, 4) = Range.Cells(Partitions(i, 2), 1).Resize(Partitions(i, 3) - Partitions(i, 2) + 1, Range.Columns.Count)
    Next i
    
    PartitionRange = Partitions2
End Function

'@Description "Returns the Selection object if it is a valid Range."
Public Function TryGetSelectionRange(ByRef OutRange As Range) As Boolean
Attribute TryGetSelectionRange.VB_Description = "Returns the Selection object if it is a valid Range."
    If Selection Is Nothing Then Exit Function
    If Not TypeOf Selection Is Range Then Exit Function
    
    Set OutRange = Selection
    TryGetSelectionRange = True
End Function
