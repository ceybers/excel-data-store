VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteValues2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("RemoteDataStore.Models.Plural")
Option Explicit

Private Const ARRAY_WIDTH As Long = 9
Private Enum RemoteValuesColumnIndexes
    RV_KEY_ID = 2
    RV_FIELD_ID = 3
    RV_VALUE = 4
    RV_LATEST_HASH = 5
    RV_MTIME = 7
End Enum

Private Type TState
    Worksheet As Worksheet
    HistoryWorksheet As Worksheet
    Partitions As Variant ' PartitionRange array (1 to n, 1 to 4)
    LatestValues As Variant ' Value2 array (1 to n, 1 to m)
    ValuesByField As Scripting.Dictionary
    IsDirty As Boolean
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.ValuesByField = New Scripting.Dictionary
End Sub

' Referenced by Remote.Values (property) (getter)
' Referenced by Remote.Reload (method)
Public Sub Load(ByVal Worksheet As Worksheet, ByVal HistoryWorksheet As Worksheet)
    Log.Message "RemoteValues.Load", "RValues", Debug_Level
    Set This.Worksheet = Worksheet
    Set This.HistoryWorksheet = HistoryWorksheet
End Sub

' Referenced by: PushQuery.Run (method)
' Referenced by: PullQuery.Run (method)
' Referenced by: PullDryRunQuery.Run (method)
Public Function GetValues(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant) As Variant
    Log.Message " RemoteValues.GetValues... start", "RValues", Info_Level
    
    Dim KeyCount As Long
    KeyCount = UBound(KeyIDs, 1)
    
    Dim FieldCount As Long
    FieldCount = UBound(FieldIDs, 1)
    
    Dim Result As Variant
    ReDim Result(1 To KeyCount, 1 To FieldCount)
    GetValues = Result
    
    Log.Message "  LazyLoadPartitions", "RValues", Verbose_Level
    LazyLoadPartitions
    
    If IsEmpty(This.Partitions) Then
        GetValues = Result
        Exit Function
    End If
    
    Log.Message "  for i = 1 to " & FieldCount, "RValues", Verbose_Level
    Log.Message "  for j = 1 to " & KeyCount, "RValues", Verbose_Level
    Dim i As Long
    For i = 1 To FieldCount
        Dim ValuesForField As Variant
        ValuesForField = GetValuesForField(FieldIDs(i))
        
        Dim j As Long
        For j = 1 To KeyCount
            Dim RowIndex As Long
            RowIndex = BinarySearch2(ValuesForField, KeyIDs(j))
            
            If RowIndex > -1 Then
                Result(j, i) = ValuesForField(RowIndex, 3)
            End If
        Next j
    Next i
    
    GetValues = Result
    Log.Message " RemoteValues.GetValues... end", "RValues", Info_Level
End Function

Private Sub LazyLoadPartitions()
    Log.Message " RemoteValues.LazyLoadPartitions... start", "RValues", Info_Level
    
    If Not IsEmpty(This.Partitions) Then Exit Sub

    ' If we split Values into a ValuesLatest and ValuesHistory table, then
    ' we can skip this and save on a Sort operation (70ms+)
    Log.Message " Worksheet Sort", "RValues", Verbose_Level
    This.Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=This.Worksheet.Columns.Item(RV_LATEST_HASH)
    
    Dim FirstLatestHashCell As Range
    Set FirstLatestHashCell = This.Worksheet.Cells.Item(2, RV_LATEST_HASH)
    
    If FirstLatestHashCell.Value2 = vbEmpty Then
        ' There are no rows in the Values worksheet with a LATEST_HASH.
        Exit Sub
    End If
    
    Log.Message " Worksheet End", "RValues", Verbose_Level
    Dim LatestHashCount As Long
    LatestHashCount = FirstLatestHashCell.End(xlDown).Row - FirstLatestHashCell.Row + 1
    
    Dim LatestHashRange As Range
    Set LatestHashRange = This.Worksheet.Cells.Item(2, 1).Resize(LatestHashCount, ARRAY_WIDTH)
        
    Log.Message " PartitionRange", "RValues", Verbose_Level
    This.Partitions = PartitionRangeByColumn(LatestHashRange, RemoteValuesColumnIndexes.RV_FIELD_ID)
    QuickSort2 This.Partitions
    
    Log.Message " Get Value2", "RValues", Error_Level
    This.LatestValues = LatestHashRange.Offset(0, 1).Resize(, 3).Value2
    
    Log.Message " RemoteValues.LazyLoadPartitions... end", "RValues", Info_Level
End Sub

Private Function GetValuesForField(ByVal FieldID As String) As Variant
    If IsEmpty(This.Partitions) Then Exit Function
    
    If This.ValuesByField.Exists(FieldID) Then
        GetValuesForField = This.ValuesByField.Item(FieldID)
        Exit Function
    End If
    
    Dim PartitionIndex As Long
    PartitionIndex = BinarySearch2(This.Partitions, FieldID)
    
    If PartitionIndex = -1 Then
        Dim EmptyArray As Variant
        ReDim EmptyArray(1 To 1, 1 To 3)
        GetValuesForField = EmptyArray
        Exit Function
    End If
    
    Dim FirstRow As Long
    FirstRow = This.Partitions(PartitionIndex, 2)
        
    Dim RowCount As Long
    RowCount = This.Partitions(PartitionIndex, 3) - FirstRow + 1
    
    Dim ThisBox As Variant
    ThisBox = ArrayBox(This.LatestValues, FirstRow, 1, RowCount, 3)
    QuickSort2 ThisBox

    This.ValuesByField.Add Key:=FieldID, Item:=ThisBox
    GetValuesForField = ThisBox
End Function

' Returns an array containing RemoteValue elements for the given combination of KeyID and FieldID.
' Returns Nothing if it cannot find any RemoteValues.
' Referenced by: ValueTimelineQuery.Run (method)
Public Function GetAllValues(ByVal KeyID As String, ByVal FieldID As String) As Variant
    Dim HistoryRange As Range
    Set HistoryRange = WorksheetHelpers.GetWorksheetDatabodyRange(This.HistoryWorksheet)
    
    Dim FieldPartitions As Variant
    FieldPartitions = PartitionRangeByColumn(HistoryRange, RV_FIELD_ID)
    If IsEmpty(FieldPartitions) Then Exit Function
    QuickSort2 FieldPartitions
    
    Dim FieldIndex As Long
    FieldIndex = BinarySearch2(FieldPartitions, FieldID)
    If FieldIndex = -1 Then Exit Function
    
    Dim KeyPartitions As Variant
    KeyPartitions = PartitionRangeByColumn(FieldPartitions(FieldIndex, 4), RV_KEY_ID)
    If IsEmpty(KeyPartitions) Then Exit Function
    QuickSort2 KeyPartitions
        
    Dim KeyIndex As Long
    KeyIndex = BinarySearch2(KeyPartitions, KeyID)
    If KeyIndex = -1 Then Exit Function
    
    Dim ValuesRange As Range
    Set ValuesRange = KeyPartitions(KeyIndex, 4)
    KeyPartitions(KeyIndex, 4).Sort Header:=xlNo, _
        Key1:=This.HistoryWorksheet.Columns(RV_MTIME), _
        Order1:=xlDescending
    
    Dim vv As Variant
    vv = ValuesRange.Value2
    
    Dim Result As Variant
    ReDim Result(1 To UBound(vv, 1))
    
    Dim i As Long
    For i = 1 To UBound(vv, 1)
        Dim NewRemoteValue As RemoteValue
        Set NewRemoteValue = New RemoteValue
        NewRemoteValue.LoadFromArray vv, i
        Set Result(i) = NewRemoteValue
    Next i
    
    GetAllValues = Result
End Function

' Referenced by: PushQuery.Run (method)
Public Sub UpsertValuesWithMask(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, _
    ByVal Values As Variant, ByVal ChangeMask As Variant, _
    ByVal CommitID As String)
    
    Dim i As Long
    For i = 1 To UBound(FieldIDs, 1)
        UpsertField i, KeyIDs, FieldIDs, Values, ChangeMask, CommitID
    Next i
    
    If This.IsDirty Then
        This.Partitions = Empty
        This.LatestValues = Empty
        This.ValuesByField.RemoveAll
        This.IsDirty = False
    End If
End Sub

Private Sub UpsertField(ByVal FieldIndex As Long, _
    ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, _
    ByVal Values As Variant, ByVal ChangeMask As Variant, _
    ByVal CommitID As String)
    
    Dim ThisFieldID As String
    ThisFieldID = FieldIDs(FieldIndex)
    
    Dim NewRemoteValues As Collection
    Set NewRemoteValues = CreateNewRemoteValues(FieldIndex, _
        KeyIDs, FieldIDs, Values, ChangeMask, CommitID)
    
    Dim ChangeCount As Long
    ChangeCount = NewRemoteValues.Count
    
    If ChangeCount = 0 Then Exit Sub
    
    Dim ExistingValuesForField As Variant
    ExistingValuesForField = GetValuesForField(ThisFieldID)
    
    Dim NoExistingValues As Boolean
    NoExistingValues = IsEmpty(ExistingValuesForField)
    
    Dim RemoteValuesToUpdate As Collection
    Set RemoteValuesToUpdate = New Collection
    Dim RemoteValuesToInsert As Collection
    Set RemoteValuesToInsert = New Collection
    
    Dim ThisRemoteValue As RemoteValue
    For Each ThisRemoteValue In NewRemoteValues
        If NoExistingValues = True Then
            RemoteValuesToInsert.Add Item:=ThisRemoteValue
        ElseIf BinarySearch2(ExistingValuesForField, ThisRemoteValue.KeyID) = -1 Then
            RemoteValuesToInsert.Add Item:=ThisRemoteValue
        Else
            RemoteValuesToUpdate.Add Item:=ThisRemoteValue
        End If
    Next ThisRemoteValue
    
    InsertRemoteValuesToWorksheet RemoteValuesToInsert
    UpdateRemoteValuesToWorksheet RemoteValuesToUpdate
    SaveHistoricalRemoteValues NewRemoteValues
    This.IsDirty = True
End Sub

Private Sub InsertRemoteValuesToWorksheet(ByVal RemoteValues As Collection)
    If RemoteValues.Count = 0 Then Exit Sub
    
    Log.Message "RemoteValues2.InsertRemoteValuesToWorksheet... start", "RValues", Info_Level
    If RemoteValues.Count = 0 Then Exit Sub
        
    Dim vv As Variant
    ReDim vv(1 To RemoteValues.Count, 1 To ARRAY_WIDTH)
    
    Dim i As Long
    For i = 1 To RemoteValues.Count
        Dim ThisRemoteValue As RemoteValue
        Set ThisRemoteValue = RemoteValues.Item(i)
        
        ThisRemoteValue.SaveToArray vv, i
        
        ThisRemoteValue.Latest = ThisRemoteValue.ID
    Next i
    
    Log.Message "RemoteValues2.RangeSetValueFromVariant... start", "RValues", Debug_Level
    RangeSetValueFromVariant This.Worksheet.Cells.Item(This.Worksheet.UsedRange.Rows.Count + 1, 1), vv
    Log.Message "RemoteValues2.RangeSetValueFromVariant... done", "RValues", Debug_Level
    
    Log.Message "RemoteValues2.InsertRemoteValuesToWorksheet... done", "RValues", Info_Level
End Sub

Private Sub UpdateRemoteValuesToWorksheet(ByVal RemoteValues As Collection)
    If RemoteValues.Count = 0 Then Exit Sub
    
    Log.Message "RemoteValues2.UpdateRemoteValuesToWorksheet... start", "RValues", Info_Level
    Dim FieldID As String
    FieldID = RemoteValues.Item(1).FieldID
    
    Dim PartitionRange As Range
    Set PartitionRange = This.Partitions(BinarySearch2(This.Partitions, FieldID), 4)
    
    Dim vv As Variant
    vv = PartitionRange.Value2
    
    Dim LookupArray As Variant
    ReDim LookupArray(1 To UBound(vv, 1), 1 To 2)
    
    Dim i As Long
    For i = 1 To UBound(vv, 1)
        LookupArray(i, 1) = vv(i, 2)
        LookupArray(i, 2) = i
    Next i
    QuickSort2 LookupArray
    
    Dim ThisRemoteValue As RemoteValue
    For Each ThisRemoteValue In RemoteValues
        Dim Row As Long
        Row = BinarySearch2(LookupArray, ThisRemoteValue.KeyID)
        
        Dim OldValueID As String
        OldValueID = vv(Row, 1)
        
        ThisRemoteValue.SaveToArray vv, Row
        
        ThisRemoteValue.Latest = OldValueID
    Next ThisRemoteValue
    
    Log.Message "RemoteValues2.RangeSetValueFromVariant... start", "RValues", Debug_Level
    RangeSetValueFromVariant PartitionRange, vv
    Log.Message "RemoteValues2.RangeSetValueFromVariant... done", "RValues", Debug_Level
    
    Log.Message "RemoteValues2.UpdateRemoteValuesToWorksheet... done", "RValues", Info_Level
End Sub

Private Sub SaveHistoricalRemoteValues(ByVal RemoteValues As Collection)
    Dim vv As Variant
    ReDim vv(1 To RemoteValues.Count, 1 To ARRAY_WIDTH)
    
    Dim i As Long
    For i = 1 To RemoteValues.Count
        Dim ThisRemoteValue As RemoteValue
        Set ThisRemoteValue = RemoteValues.Item(i)
        ThisRemoteValue.SaveToArray vv, i
    Next i
    
    RangeSetValueFromVariant This.HistoryWorksheet.Cells.Item(This.HistoryWorksheet.UsedRange.Rows.Count + 1, 1), vv
End Sub

Private Function CreateNewRemoteValues(ByVal FieldIndex As Long, _
    ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, _
    ByVal Values As Variant, ByVal ChangeMask As Variant, _
    ByVal CommitID As String) As Collection

    Dim Result As Collection
    Set Result = New Collection
    
    Dim ThisFieldID As String
    ThisFieldID = FieldIDs(FieldIndex)
    
    Dim i As Long
    For i = 1 To UBound(KeyIDs, 1)
        If ChangeMask(i, FieldIndex) = True Then
            Dim NewRemoteValue As RemoteValue
            
            Set NewRemoteValue = CreateNewValue(KeyIDs(i), _
                ThisFieldID, _
                Values(i, FieldIndex), _
                CommitID)
                
            Result.Add NewRemoteValue
        End If
    Next i
    
    Set CreateNewRemoteValues = Result
End Function

Private Function CreateNewValue(ByVal KeyID As String, ByVal FieldID As String, _
    ByVal Value As Variant, ByVal CommitID As String) As RemoteValue
    Dim NewRemoteValue As RemoteValue
    Set NewRemoteValue = New RemoteValue
    With NewRemoteValue
        .KeyID = KeyID
        .FieldID = FieldID
        .Value = Value
        .Latest = KeyID & "x" & FieldID
        .CreationTime = Application.Evaluate("=NOW()")
        .ModificationTime = .CreationTime
        .CommitID = CommitID
        .ID = Left$(HashSHA1(.Latest & .CreationTime), HASH_LEN)
    End With
    Set CreateNewValue = NewRemoteValue
End Function

