Attribute VB_Name = "modRebuildRemoteFields"
'@Folder "RemoteDataStore.Factories"
Option Explicit

Private Const COL_ID As Long = 1
Private Const COL_PATH As Long = 2
Private Const COL_KEY As Long = 3

Public Sub RebuildRemoteTable(ByVal Worksheet As Worksheet, ByVal ColumnHeadings As Variant)
    RebuildHeaders Worksheet, ColumnHeadings

    RemoveExtraColumns Worksheet, ColumnHeadings
    
    RemoveExtraRows Worksheet
    
    RemoveDuplicates Worksheet

    RebuildIDs Worksheet
End Sub

Private Sub RebuildHeaders(ByVal Worksheet As Worksheet, ByVal ColumnHeadings As Variant)
    ' TODO FIX ArrayTransform.ArrayToRow expects a 1-based array not a 0-based array.
    RangeSetValueFromVariant Worksheet.Cells.Item(1, 1), ArrayTransform.ArrayToRow(ColumnHeadings)
End Sub

Private Sub RemoveExtraColumns(ByVal Worksheet As Worksheet, ByVal ColumnHeadings As Variant)
    Dim ColumnCount As Long
    ColumnCount = UBound(ColumnHeadings) + 1
    
    Dim LastCell As Range
    Set LastCell = Worksheet.UsedRange.Cells.Item(Worksheet.UsedRange.Cells.Count)
    
    If LastCell.Column > ColumnCount Then
        Dim ExtraColumnsToDelete As Range
        Set ExtraColumnsToDelete = Worksheet.Range(Worksheet.Cells.Item(1, ColumnCount + 1), LastCell).EntireColumn
        ExtraColumnsToDelete.Delete
    End If
End Sub

Private Sub RemoveExtraRows(ByVal Worksheet As Worksheet)
    Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=Worksheet.Columns.Item(COL_KEY), _
        Order1:=xlDescending
        
    Dim ActualUsedRange As Range
    Set ActualUsedRange = Worksheet.Cells.Columns.Item(COL_KEY).SpecialCells(xlCellTypeConstants)
    
    Dim LastActualCell As Range
    Set LastActualCell = ActualUsedRange.Cells.Item(ActualUsedRange.Cells.Count)
    
    Dim LastCell As Range
    Set LastCell = Worksheet.UsedRange.Cells.Item(Worksheet.UsedRange.Cells.Count)
    
    If LastActualCell.Row < LastCell.Row Then
        Dim ExtraRowsToDelete As Range
        Set ExtraRowsToDelete = Worksheet.Range(LastActualCell.Offset(1, 0), LastCell).EntireRow
        ExtraRowsToDelete.Delete
    End If
End Sub

Private Sub RemoveDuplicates(ByVal Worksheet As Worksheet)
    Worksheet.UsedRange.RemoveDuplicates Header:=xlYes, Columns:=Array(COL_PATH, COL_KEY)
End Sub

Private Sub RebuildIDs(ByVal Worksheet As Worksheet)
    Dim DataBodyRange As Range
    Set DataBodyRange = Worksheet.Cells.Columns.Item(COL_KEY).SpecialCells(xlCellTypeConstants)
    
    Dim Range As Range
    Set Range = Worksheet.Range(Worksheet.Cells.Item(2, 1), DataBodyRange.Cells.Item(DataBodyRange.Cells.Count))
    
    Dim vv As Variant
    vv = Range.Value2
    
    Dim i As Long
    For i = 1 To UBound(vv, 1)
        vv(i, COL_ID) = HashSHA1(vv(i, COL_PATH) & "\" & vv(i, COL_KEY))
    Next i
    
    RangeSetValueFromVariant Range, vv
    
    Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=Worksheet.Columns.Item(COL_ID), _
        Order1:=xlAscending
End Sub
