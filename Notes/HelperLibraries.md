# Helper Libraries
## CollectionHelpers
### CollectionClear
Removes all the items in a `Collection`.
```vb
CollectionClear(ByVal Collection As Collection)
```

### CollectionToRangeUnion
Returns a `Range` which is the Union of all the `Range` items in a `Collection`. Returns `Nothing` if the `Collection` is `Nothing` or if the `Collection` is empty.

```vb
CollectionToRangeUnion(ByVal Collection As Collection) As Range
```

## ListColumnHelpers
### TryGetListColumn
Returns `True` if a `ListColumn` with the specified name exists in the `ListObject` and sets the provided variable to the `ListColumn` object. Returns `False` if it does not.

```vb
TryGetListColumn(
    ByVal ListColumnName As String, 
    ByVal ListObject As ListObject, 
    ByRef OutListColumn As ListColumn) As Boolean
```

### ListColumnExists
Returns `True` if a `ListColumn` with the specified name exists in the `ListObject`. Returns `False` if it does not.

```vb
ListColumnExists(
    ByVal ListColumnName As String, 
    ByVal ListObject As ListObject) As Boolean
```

## ListObjectHelpers
### GetAllListObjects
Returns a `Collection` containing all the `ListObject`s in the specified `Workbook`. The `Name` property of each `ListObject` is used as the Key in the `Collection`. Returns an empty `Collection` if there are no `ListObject`s.

```vb
GetAllListObjects(
    ByVal Workbook as Workbook) As Collection
```

### TryGetListObjectFromCollection
Returns `True` if a `ListObject` with the specified `Name` exists in a `Collection` of `ListObjects`, and sets the provided variable to the `ListObject` object. Returns `False` if nothing is found.

```vb
TryGetListObjectFromCollection(
    ByVal ListObjectName As String, 
    ByVal TableCollection As Collection, 
    ByRef OutListObject As ListObject) As Boolean
```

### IsListObjectProtected
Returns `True` if any of the cells in the `ListObject`'s `Range` are Locked *and* the `Worksheet` is Protected, or if the `Workbook` is opened in Protected Viewing mode.

```vb
IsListObjectProtected(
    ByVal ListObject As ListObject) As Boolean
```

### TryGetSelectedListObject
Returns `True` if the `Selection` object contains a `ListObject` and sets the provided variable to the `ListObject`. Returns `False` if nothing is found.

```vb
TryGetSelectedListObject(
    ByRef OutListObject As ListObject) As Boolean
```

### TryGetSingleListObjectInActiveSheet
Returns `True` if there is exactly one `ListObject` in the `ActiveSheet` and sets the provided variable to the `ListObject`. Returns `False` if there are zero or more than one `ListObject`s.

```vb
TryGetSingleListObjectInActiveSheet(
    ByRef OutListObject As ListObject) As Boolean
```

### TryGetListObjectInWorksheet
Returns `True` if the `Worksheet` contains a `ListObject` with the specified name and sets the provided variable to the `ListObject` object. Returns `False` if nothing is found.

```vb
TryGetListObjectInWorksheet(
    ByVal ListObjectName As String, 
    ByVal Worksheet As Worksheet, 
    ByRef OutListObject As ListObject) As Boolean
```

## RangeHelpers
### ResizeRangeToArray
Returns a new `Range` with the same shape as the specified 2-dimensional array, starting, from the top-left cell in the  specified `Range`. Throws an error if the input `Range` is `Nothing` or if the array is not 2-dimensional.

```vb
ResizeRangeToArray(
    ByVal InputRange As Range, 
    ByVal InputArray As Variant) As Range
```

### RangeSetValueFromVariant
 Updates the `Value2` property of the cells in a `Range` with the values from a 2-dimensional Variant array. If the array is smaller than the `Range`, only the cells from the top-left to the extents of the array will be updated. If the `Range` is larger than the `Range`, the function will update cells outside of the given `Range`.

```vb
RangeSetValueFromVariant(
    ByVal InputRange As Range, 
    ByVal InputVariantArray As Variant)
```

### RangeBox
Returns a new `Range` offset and resized from specified input `Range`. Returns `Nothing` if the input `Range` is nothing. Throws an error if any of the indices are zero or negative.

```vb
RangeBox(
    ByVal InputRange As Range, 
    ByVal Row As Long, ByVal Column As Long, _
    ByVal Rows As Long, ByVal Columns As Long) As Range
```
### PartitionRange
See [PartitionRange](Partition.md).

### TryGetSelectionRange
Returns `True` if the `Selection` object is of type `Range` and sets the variable to the `Range` object. Returns `False` if `Selection` is `Nothing` or is not a `Range`.

```vb
TryGetSelectionRange(ByRef OutRange As Range) As Boolean
```

### TryIntersectRanges
Returns `True` if the two specified `Range`s can be intersected and sets the output variable to the intersected `Range`. Returns `False` if they cannot be intersected or if one or both of the `Range` are `Nothing`.

```vb
TryIntersectRanges(
    ByVal Range1 As Range, 
    ByVal Range2 As Range, 
    ByRef OutRange As Range) As Boolean
```

### GetStaggeredArrayValues
Returns the `Value2` property array of a non-contiguous `Range` that has multiple `Area`s. The output `Variant` array is of the same shape as the `BaseRange` parameter. Cells that are not in the `SelectedRange` will be `Empty` variants.

```vb
GetStaggeredArrayValues(
    ByVal BaseRange As Range, 
    ByVal SelectedRange As Range) As Variant
```

## WorkbookHelpers
### TryGetWorkbook
Returns `True` if a `Workbook` with the specified name is open in Excel and sets the provided variable to the `Workbook` object. Returns `False` if nothing is found.

```vb
TryGetWorkbook(
    ByVal WorkbookName As String, 
    ByRef OutWorkbook As Workbook) As Boolean
```

### IsWorkbookProtectedView
Returns True if the specified `Workbook` is opened in Protected View. Returns `False` if it is not, or if the variable is set to `Nothing`.

```vb
IsWorkbookProtectedView(ByVal WorkbookName As String) As Boolean
```

### IsWorkbookOpen
Returns `True` if the specified variable is referencing a `Workbook` that is still open. Returns `False` if the variable is referencing a `Workbook` that has been closed, or if the variable is set to `Nothing`.

```vb
IsWorkbookOpen(ByVal Workbook As Workbook) As Boolean
```

## WorksheetHelpers
### GetWorksheetDatabodyRange
Returns the range of cells in a `Worksheet` starting from the first row beneath the header row ranging until the last row in the `UsedRange`. Assumes header row is always Row 1. Returns `Nothing` if there are no rows, or if there is only a header row.

```vb
GetWorksheetDatabodyRange(ByVal Worksheet As Worksheet) As Range
```
### TryGetWorksheet
Returns `True` if the `Worksheet` with the specified name exists in the specified `Workbook` and sets the provided variable to the `Worksheet` object. Returns `False` if nothing is found.
```vb
TryGetWorksheet(
    ByVal WorksheetName As String, 
    ByVal Workbook As Workbook, _
    ByRef OutWorksheet As Worksheet) As Boolean
```

---