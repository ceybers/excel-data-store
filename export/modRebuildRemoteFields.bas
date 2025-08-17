Attribute VB_Name = "modRebuildRemoteFields"
'@IgnoreModule IndexedDefaultMemberAccess
'@Folder "RemoteDataStore.Factories"
Option Explicit

Private Const COL_ID As Long = 1
Private Const COL_PATH As Long = 2
Private Const COL_KEY As Long = 3

Public Sub RebuildRemoteTable(ByVal Worksheet As Worksheet, ByVal ColumnHeadings As Variant)
    RebuildHeaders Worksheet, ColumnHeadings

    RemoveExtraColumns Worksheet, UBound(ColumnHeadings) + 1
    
    RemoveExtraRows Worksheet
    
    Worksheet.UsedRange.RemoveDuplicates Header:=xlYes, _
        Columns:=Array(COL_PATH, COL_KEY)
        
    Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=Worksheet.Columns.Item(COL_KEY), Order1:=xlDescending

    RebuildIDs Worksheet
End Sub

Private Sub RebuildHeaders(ByVal Worksheet As Worksheet, ByVal ColumnHeadings As Variant)
    RangeSetValueFromVariant Worksheet.Cells(1, 1).Resize(1, UBound(ColumnHeadings) + 1), ArrayTransform.ArrayToRow(ColumnHeadings)
    
    Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=Worksheet.Columns.Item(COL_KEY), Order1:=xlDescending
End Sub

Private Sub RemoveExtraColumns(ByVal Worksheet As Worksheet, ByVal ColumnCount As Long)
    Dim LastCell As Range
    Set LastCell = Worksheet.UsedRange.Cells(Worksheet.UsedRange.Cells.Count)
    
    If LastCell.Column > ColumnCount Then
        Dim ExtraColumnsToDelete As Range
        Set ExtraColumnsToDelete = Worksheet.Range(Worksheet.Cells(1, ColumnCount + 1), LastCell).EntireColumn
        ExtraColumnsToDelete.Delete
    End If
End Sub

Private Sub RemoveExtraRows(ByVal Worksheet As Worksheet)
    Dim ActualUsedRange As Range
    Set ActualUsedRange = Worksheet.Cells.Columns.Item(COL_KEY).SpecialCells(xlCellTypeConstants)
    
    Dim LastActualCell As Range
    Set LastActualCell = ActualUsedRange.Cells(ActualUsedRange.Cells.Count)
    
    Dim LastCell As Range
    Set LastCell = Worksheet.UsedRange.Cells(Worksheet.UsedRange.Cells.Count)
    
    If LastActualCell.Row < LastCell.Row Then
        Dim ExtraRowsToDelete As Range
        Set ExtraRowsToDelete = Worksheet.Range(LastActualCell.Offset(1, 0), LastCell).EntireRow
        ExtraRowsToDelete.Delete
    End If
End Sub

Private Sub RebuildIDs(ByVal Worksheet As Worksheet)
    Dim DataBodyRange As Range
    Set DataBodyRange = Worksheet.Cells.Columns.Item(COL_KEY).SpecialCells(xlCellTypeConstants)
    
    Dim Range As Range
    Set Range = Worksheet.Range(Worksheet.Cells(2, 1), DataBodyRange.Cells(DataBodyRange.Cells.Count))
    
    Dim vv As Variant
    vv = Range.Value2
    
    Dim i As Long
    For i = 1 To UBound(vv, 1)
        vv(i, COL_ID) = HashSHA1(vv(i, COL_PATH) & "\" & vv(i, COL_KEY))
    Next i
    
    RangeSetValueFromVariant Range, vv
    
    Worksheet.UsedRange.Sort Header:=xlYes, Key1:=Worksheet.Columns.Item(COL_ID), Order1:=xlAscending
End Sub
