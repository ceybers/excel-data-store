Attribute VB_Name = "ArrayCheck"
'@IgnoreModule ProcedureNotUsed
'@Folder("Helpers.Array")
Option Explicit

' Returns True if InputArray is a 2-dimensional array with one-based indexes.
Public Function IsTwoDimensionalOneBasedArray(ByVal InputArray As Variant) As Boolean
    If Not IsArray(InputArray) Then Exit Function
    If LBound(InputArray, 1) <> 1 Then Exit Function
    On Error GoTo ErrorNotTwoDimensional
    If LBound(InputArray, 2) <> 1 Then Exit Function
    On Error GoTo 0
    
    IsTwoDimensionalOneBasedArray = True
    Exit Function
    
ErrorNotTwoDimensional:
    If Err.Number = 9 Then Exit Function
End Function

' Returns True if InputArray is a 1-dimensional array with one-based indexes.
Public Function IsOneDimensionalOneBasedArray(ByVal InputArray As Variant) As Boolean
    If Not IsArray(InputArray) Then Exit Function
    If LBound(InputArray, 1) <> 1 Then Exit Function
    On Error GoTo OkayNotTwoDimensional
    If LBound(InputArray, 2) <> 1 Then Exit Function
    On Error GoTo 0
    Exit Function
    
OkayNotTwoDimensional:
    IsOneDimensionalOneBasedArray = True
End Function
