# PartitionRangeByColumn() method
Partitions a Range into multiple smaller Ranges based on the values in a specific column. Each unique value in this Partitioning Column will result in a partitioned Range that contains all the rows with that value.

## Syntax
RangePartition.**PartitionRangeByColumn**(Range, Column) As Variant

## Parameters
| Name | Required/Optional | Data Type | Description |
| ---- | ----------------- | --------- | ----------- |
| _Range_ | Required | Range | The range that will be paritioned. |
| _Column_ | Required | Long | The index of the column that will be used to partition on. (1-based) | 

## Return value
Returns a 2-dimensional array. Each row is a unique value from the partitioning column. The array will have 4 columns as follows:
- `Result(i, 1)` is the Value in the partitioning column.
- `Result(i, 2)` is the index of the first row where this value appears in the input Range.
- `Result(i, 3)` is the index of the last row where this value appears in the input Range.
- `Result(i, 4)` is the Range of the partition.

## Remarks
> [!CAUTION]
> This method will Sort rows in the Worksheet. Existing sort order will not necessarily be retained. It will fail if it cannot perform the Sort operation.

- The input range must be one non-contiguous area. (No hidden or filtered rows).
- The input range must not include a Header Row.
- The input range must not contain a ListObject (table).
- The index of the first and last row are relative to the top-left cell in the Range.
- Values in the Partioning Column may be any valid Excel Variant Type except `vbError`.
- Values of type `Error` will be replaced with `Empty`.

## Examples

|     | A       | B       | C	  | D   |
| --- | ------- | ------- | ----- | --- |
|  1  | **Level A** | **Level B** | **Data**	| **Rnd** |
|  2  | alpha   | apple   | $C$2  | 765 |
|  3  | alpha   | apple   | $C$3  | 550 |
|  4  | alpha   | banana  | $C$4  | 628 |
|  5  | bravo   | banana  | $C$5  | 981 |
|  6  | bravo   | carrot  | $C$6  | 426 |
|  7  | bravo   | carrot  | $C$7  | 300 |
|  8  | charlie | éclair  | $C$10 | 975 |
|  9  | charlie | donut   | $C$8  | 127 |
| 10  | charlie | donut   | $C$9  | 937 |

This example partitions the sample table on the first column. It returns an array containing 3 rows, one for each value in the partition column (alpha, bravo, charlie).

```vb
Dim InputRange As Range
Set InputRange = ActiveSheet.Range("A2:D10")

Dim Partitions As Variant
Partitions = PartitionRangeByColumn(InputRange, 1)

' Partitions(2, 1) = "bravo"
' Partitions(2, 2) = 4
' Partitions(2, 3) = 6
' Partitions(2, 4) = Range("A5:D7")
```

This example further partitions one of the partitions from the previous example, this time partitioning the range on the second column. It returns an array containing 2 rows, one for each value in the second partition column (banana, carrot).

```vb
Dim SubPartitions As Variant
SubPartitions = PartitionRangeByColumn(Partitions(2, 4), 2)

' SubPartitions(1, 1) = "banana"
' SubPartitions(1, 2) = 4
' SubPartitions(1, 3) = 4
' SubPartitions(1, 4) = Range("A5:D5")
```

## Related
Because the result value returns an array with the Partition Value in the first column, it is possible to use `BinarySearch2` to search for values in the array. However, `BinarySearch2` will require the resulting array to be sorted with `QuickSort2` and not the native Excel `Sort` method.

```vb
ArraySort.QuickSort2 Partitions

Dim Index As Variant
Index = ArraySearch.BinarySearch2(Partitions, "banana")
' Index = 2

Index = ArraySearch.BinarySearch2(Partitions, "foobar")
' Index = -1
```

# See Also
- [Helper Libraries documentation](HelperLibraries.md)