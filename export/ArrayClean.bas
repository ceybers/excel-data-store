Attribute VB_Name = "ArrayClean"
'@Folder("Helpers.Array")
Option Explicit

'@Description "Replaces all the cells of VarType vbError with cells of VarType vbEmpty in a 2-dimensional array."
Public Sub ReplaceErrorCells(ByRef InputArray As Variant)
Attribute ReplaceErrorCells.VB_Description = "Replaces all the cells of VarType vbError with cells of VarType vbEmpty in a 2-dimensional array."
    If Not IsArray(InputArray) Then Exit Sub
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
