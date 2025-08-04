# TableToAreas
## Getting a Table Map from a Selection intersecting a ListObject
```vb
Public Sub DoTableMapSelectionToRange()
    Dim SomeMap As TableMap
    Set SomeMap = New TableMap
    
    ' Creates a TableMap, storing the intersecting Key and Field indexes in memory
    SomeMap.LoadFromListObjectSelection Range("F8:M14").SpecialCells(xlCellTypeVisible), "Key"

    ' Saves the TableMap template (ListObject Name, Key Column Name, Key Values and Indexes, Field Names and Indexes) in the shape of a table, but without the values
    SomeMap.SaveToRange Range("A33")

    ' Saves the values into the range. Doesn't add the field and key headings if they are not present.
    SomeMap.SaveValuesToRange Range("A33")
End Sub
```

## Getting a Table Map from data at runtime
```vb
Public Sub DoTableMapCodeToRange()
    Dim SomeMap As TableMap
    Set SomeMap = New TableMap
    
    Dim ListObject As ListObject
    Set ListObject = ActiveSheet.ListObjects.Item(1)
    
    Dim KeyColumnName As String
    KeyColumnName = "Key"
    
    Dim FieldColumnNames As Variant
    FieldColumnNames = Array("Field3", "Field4", "Field7")
    
    SomeMap.LoadFromListObject ListObject, KeyColumnName, FieldColumnNames
End Sub
```

## Updating a ListObject from a TableMap stored in a Range
```vb
Public Sub DoTableMapRangeToListObject()
    Dim SomeMap As TableMap
    Set SomeMap = New TableMap

    ' Loads a TableMap from a Range. Tries to resolve the ListObject from the name.
    ' Loads the key and field indexes into memory.
    SomeMap.LoadFromRange Range("A33")

    ' Uses the TableMap to update the ListObject using the mapping in the TableMap.
    ' The actual values are sourced from the Range. Only the indexes are used from the TableMap
    ' object. Returns an error if any of the key indexes or field indexes do not match what key
    ' or fields are in the ListObject at that location.
    SomeMap.UpdateListObjectFromRange Range("A33")
End Sub
```