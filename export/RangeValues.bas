Attribute VB_Name = "RangeValues"
'@Folder("Helpers.Range")
Option Explicit

'@Description "Updates the Value2 property of the cells in a Range with the values from a 2-dimensional Variant array. If the array is smaller than the Range, only the cells from the top-left to the extents of the array will be updated. If the Range is larger than the Range, the function will update cells outside of the given Range."
Public Sub RangeSetValueFromVariant(ByVal InputRange As Range, ByVal InputVariant As Variant)
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputVariant)
    
    Log.Message "RangeSetValueFromVariant writing to = " & InputRange.Address(False, False) & " with Variant(" & UBound(InputVariant, 1) & " to " & UBound(InputVariant, 2) & ")"
    InputRange.Cells.Item(1, 1).Resize(UBound(InputVariant, 1), UBound(InputVariant, 2)).Value2 = InputVariant
End Sub

' Used by MappedTable when Partially Selecting multiple areas.
'@Description "Returns the Value2 property array of a non-contiguous Range that has multiple Areas. The output Variant array is of the same shape as the BaseRange parameter. Cells that are not in the SelectedRange will be Empty variants."
Public Function GetStaggeredArrayValues(ByVal BaseRange As Range, ByVal SelectedRange As Range) As Variant
Attribute GetStaggeredArrayValues.VB_Description = "Returns the Value2 property array of a non-contiguous Range that has multiple Areas. The output Variant array is of the same shape as the BaseRange parameter. Cells that are not in the SelectedRange will be Empty variants."
    If BaseRange Is Nothing Then Exit Function
    If BaseRange.Cells.Count <= 1 Then Exit Function
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

'@Description "Updates the values of all the cells in a column Range with the values from a Variant array, respecting any filters and hidden columns."
Public Sub UpdateFilteredColumnRangeWithValues(ByVal BaseRange As Range, ByVal VariantValues As Variant)
Attribute UpdateFilteredColumnRangeWithValues.VB_Description = "Updates the values of all the cells in a column Range with the values from a Variant array, respecting any filters and hidden columns."
    Debug.Assert Not BaseRange Is Nothing
    Debug.Assert IsArray(VariantValues)
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(VariantValues)
    Debug.Assert BaseRange.Rows.Count = UBound(VariantValues, 1)
    Debug.Assert BaseRange.Columns.Count = UBound(VariantValues, 2)
    Debug.Assert BaseRange.Columns.Count = 1
    
    Dim VisibleRange As Range
    On Error Resume Next
    Set VisibleRange = BaseRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    ' Replace with TryGetSpecialCells...
    If VisibleRange Is Nothing Then
        ' Entire range is hidden
        BaseRange.Value2 = VariantValues
        Exit Sub
    End If
    
    If VisibleRange.Areas.Count = 1 Then
        If VisibleRange.Rows.Count = BaseRange.Rows.Count Then
            ' Entire range is visible
            BaseRange.Value2 = VariantValues
            Exit Sub
        End If
    End If
    
    Dim FirstRow As Long
    FirstRow = BaseRange.Row
    
    Dim OptimisticCount As Long
    OptimisticCount = (VisibleRange.Areas.Count * 2) + 1
    
    Dim AreaExtents As Variant
    ReDim AreaExtents(1 To OptimisticCount, 1 To 2)
    ' First column is the first row of the range
    ' Second column is the count of rows in that range
    
    Dim Cursor As Long
    
    Dim LeadingHiddenRangeExists As Boolean
    LeadingHiddenRangeExists = VisibleRange.Areas.Item(1).Row <> FirstRow
    
    If LeadingHiddenRangeExists Then
        Cursor = Cursor + 1
        AreaExtents(Cursor, 1) = FirstRow
        AreaExtents(Cursor, 2) = VisibleRange.Areas.Item(1).Row - FirstRow
    End If
    
    ' There is always at least 1 Area
    Cursor = Cursor + 1
    AreaExtents(Cursor, 1) = VisibleRange.Areas.Item(1).Row
    AreaExtents(Cursor, 2) = VisibleRange.Areas.Item(1).Rows.Count
    
    ' From Area #2 onwards, each Area is always preceded by a hidden row
    Dim i As Long
    For i = 2 To VisibleRange.Areas.Count
        Cursor = Cursor + 1
        AreaExtents(Cursor, 1) = AreaExtents(Cursor - 1, 1) + AreaExtents(Cursor - 1, 2)
        AreaExtents(Cursor, 2) = VisibleRange.Areas.Item(i).Row - AreaExtents(Cursor, 1)
        Cursor = Cursor + 1
        AreaExtents(Cursor, 1) = VisibleRange.Areas.Item(i).Row
        AreaExtents(Cursor, 2) = VisibleRange.Areas.Item(i).Rows.Count
    Next i
    
    Dim LastCell As Range
    Set LastCell = BaseRange.Cells.Item(BaseRange.Rows.Count, 1)
    
    ' Check for trailing hidden range
    If LastCell.EntireRow.Hidden = True Then
        Cursor = Cursor + 1
        AreaExtents(Cursor, 1) = AreaExtents(Cursor - 1, 1) + AreaExtents(Cursor - 1, 2)
        AreaExtents(Cursor, 2) = LastCell.Row + 1 - AreaExtents(Cursor, 1)
    End If
    
    ' Change column 1 from starting row to offset from the top-most row in the base range
    For i = 1 To Cursor
        AreaExtents(i, 1) = AreaExtents(i, 1) - FirstRow
    Next i
    
    Dim ExistingValues As Variant
    ExistingValues = BaseRange.Value2
    
    For i = 1 To Cursor
        Dim SubVariantValues As Variant
        ReDim SubVariantValues(1 To AreaExtents(i, 2), 1 To 1)
        
        Dim IsSubRangeChanged As Boolean
        IsSubRangeChanged = False
        
        Dim j As Long
        For j = 1 To AreaExtents(i, 2)
            SubVariantValues(j, 1) = VariantValues(AreaExtents(i, 1) + j, 1)
            If VariantValues(AreaExtents(i, 1) + j, 1) <> ExistingValues(AreaExtents(i, 1) + j, 1) Then
                IsSubRangeChanged = True
                Exit For
            End If
        Next j
        
        If IsSubRangeChanged Then
            For j = 1 To AreaExtents(i, 2)
                SubVariantValues(j, 1) = VariantValues(AreaExtents(i, 1) + j, 1)
            Next j
            
            Dim SubRange As Range
            Set SubRange = BaseRange.Resize(AreaExtents(i, 2)).Rows.Offset(AreaExtents(i, 1))
            
            SubRange.Value2 = SubVariantValues
        End If
    Next i
End Sub
