Attribute VB_Name = "ArrayUnique"
'@Folder("Helpers.Array")
Option Explicit

' Returns a 1-dimensional array that contains all the unique values in the
' InputArray. Not particularly performant. Lacks error handling for edge cases.
Public Function Unique(ByVal InputArray As Variant) As Variant
    If Not ArrayCheck.IsOneDimensionalOneBasedArray(InputArray) Then Exit Function
    
    Dim SortedArray As Variant
    SortedArray = InputArray
    ArraySort.QuickSort SortedArray
    
    Dim OutputArray As Variant
    ReDim OutputArray(LBound(SortedArray) To UBound(SortedArray))
    
    Dim C As Long
    C = LBound(SortedArray)
    
    OutputArray(C) = SortedArray(LBound(SortedArray))
    C = C + 1
    
    Dim i As Long
    For i = LBound(SortedArray) + 1 To UBound(SortedArray)
        If SortedArray(i) <> SortedArray(i - 1) Then
            OutputArray(C) = SortedArray(i)
            C = C + 1
        End If
    Next i
    
    ReDim Preserve OutputArray(LBound(OutputArray) To (C - 1))
    
    Unique = OutputArray
End Function
