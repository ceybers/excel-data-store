Attribute VB_Name = "RangePartition"
'@Folder("Helpers.Range")
Option Explicit

' Partitions a Range based on the Values in the given Column.
' CAUTION: Function will Sort rows in the Worksheet.
' Returns a 2-dimensional array with the following:
'  Result(i, 1) is the Value in the Column used to partition
'  Result(i, 2) is the first row this value appears
'  Result(i, 3) is the last row this value appears
'  Result(i, 4) is the Range of the partition
'  Error values in the Partitioning Column will be replaced with Empty.
Public Function PartitionRangeByColumn(ByVal Range As Range, ByVal Column As Long) As Variant
    If Range Is Nothing Then Exit Function
    If Range.Areas.Count <> 1 Then Exit Function
    If Range.Rows.Count = 1 Then Exit Function
    If Not Range.ListObject Is Nothing Then Exit Function
    
    If Column < 1 Then Exit Function
    If Column > Range.Columns.Count Then Exit Function
    
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
    
    Dim Cursor As Long
    Cursor = 1
    
    Partitions(Cursor, 1) = vv(1, 1)
    Partitions(Cursor, 2) = 1
    
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
        Set Partitions2(i, 4) = Range.Cells.Item(Partitions(i, 2), 1).Resize(Partitions(i, 3) - Partitions(i, 2) + 1, Range.Columns.Count)
    Next i
    
    PartitionRangeByColumn = Partitions2
End Function
