Attribute VB_Name = "ArrayClean"
'@Folder("Helpers.Array")
Option Explicit

'@Description "Replaces all the cells of VarType vbError with cells of VarType vbEmpty in a 2-dimensional array."
Public Sub ReplaceErrorCells2(ByRef InputArray As Variant)
    If Not IsTwoDimensionalOneBasedArray(InputArray) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(InputArray, 1)
        Dim j As Long
        For j = 1 To UBound(InputArray, 2)
            If IsError(InputArray(i, j)) Then
                InputArray(i, j) = Empty
            End If
        Next j
    Next i
End Sub

'@Description "Replaces all the cells of VarType vbError with cells of VarType vbEmpty in a 1-dimensional array."
Public Sub ReplaceErrorCells(ByRef InputArray As Variant)
    If Not IsOneDimensionalOneBasedArray(InputArray) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(InputArray)
        If IsError(InputArray(i)) Then
            InputArray(i) = Empty
        End If
    Next i
End Sub
