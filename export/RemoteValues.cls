VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "RemoteDataStore.Models.Plural"
Option Explicit

Private Const ARRAY_WIDTH As Long = 9
Private Const COL_KEY_ID As Long = 2
Private Const COL_FIELD_ID As Long = 3
Private Const COL_VALUE As Long = 4
Private Const COL_LATEST_HASH As Long = 5
Private Const COL_MTIME As Long = 7

Private Type TState
    Worksheet As Worksheet
    Item As Collection
    LatestHash As Variant
    IsDirty As Boolean
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Item = New Collection
End Sub

Public Property Get Count() As Long
    Count = This.Item.Count
End Property

Public Property Get Item(ByVal Index As Long) As RemoteValue
    Set Item = This.Item.Item(Index)
End Property

Public Sub Load(ByVal Worksheet As Worksheet)
    Set This.Worksheet = Worksheet
    
    RebuildCache
End Sub

Public Sub RebuildCache()
    SortWorksheet
    LoadLatestFromWorksheet
End Sub

Private Sub SortWorksheet()
    This.Worksheet.AutoFilterMode = False
    This.Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=This.Worksheet.Columns.Item(COL_LATEST_HASH)
End Sub

Private Sub LoadLatestFromWorksheet()
    CollectionClear This.Item
    This.LatestHash = Empty
    
    Dim DataBodyRange As Range
    Set DataBodyRange = GetWorksheetDatabodyRange(This.Worksheet)
    If DataBodyRange Is Nothing Then Exit Sub
    
    Dim LatestHashRange As Range
    Set LatestHashRange = This.Worksheet.Columns.Item(COL_LATEST_HASH).SpecialCells(xlCellTypeConstants)
    If LatestHashRange.Rows.Count = 1 Then Exit Sub
    
    This.LatestHash = ArrayTransform.ColumnToArray(LatestHashRange.Value2)
    
    Log.Message "RemoteValues.LoadFromWorksheet taking intersect of " & DataBodyRange.Address(False, False) & " and " & LatestHashRange.Address(False, False), "RValues", Info_Level
    If Not TryIntersectRanges(DataBodyRange, LatestHashRange.EntireRow, DataBodyRange) Then Exit Sub
    
    Dim WorksheetValues As Variant
    Log.Message "RemoteValues.LoadFromWorksheet reading from " & DataBodyRange.Address(False, False), "RValues", Info_Level
    WorksheetValues = DataBodyRange.Value2
    
    Dim i As Long
    For i = 1 To UBound(WorksheetValues, 1)
        Dim NewRemoteValue As RemoteValue
        Set NewRemoteValue = New RemoteValue
        NewRemoteValue.LoadFromArray WorksheetValues, i
        This.Item.Add Item:=NewRemoteValue
    Next i
End Sub

Public Sub Save()
    If This.IsDirty = False Then Exit Sub
    
    ClearWorksheet
    
    SetNumberFormats
    
    Dim UpdatedValues As Variant
    ReDim UpdatedValues(1 To Count, 1 To ARRAY_WIDTH)
    
    Dim i As Long
    For i = 1 To Count
        Dim ThisRemoteValue As RemoteValue
        Set ThisRemoteValue = This.Item.Item(i)
        ThisRemoteValue.SaveToArray UpdatedValues, i
    Next i
    
    If This.Worksheet.UsedRange.Rows.Count = 1 Then
        RangeSetValueFromVariant This.Worksheet.Cells.Item(2, 1), UpdatedValues
    Else
        RangeSetValueFromVariant This.Worksheet.Cells.Item(1, 1).End(xlDown).Offset(1, 0), UpdatedValues
    End If
    
    SortWorksheet
    
    RebuildCache
End Sub

Private Sub ClearWorksheet()
    This.Worksheet.AutoFilterMode = False
    If This.Worksheet.UsedRange Is Nothing Then Exit Sub
    If This.Worksheet.UsedRange.Rows.Count = 1 Then Exit Sub
    
    This.Worksheet.UsedRange.AutoFilter Field:=COL_LATEST_HASH, Criteria1:="<>"
    
    Dim FilteredDataBodyRange As Range
    
    On Error Resume Next
    Set FilteredDataBodyRange = GetWorksheetDatabodyRange(This.Worksheet).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not FilteredDataBodyRange Is Nothing Then
        FilteredDataBodyRange.EntireRow.Delete
    End If
    
    This.Worksheet.AutoFilterMode = False
End Sub

Private Sub SetNumberFormats()
    With This.Worksheet.Columns
        .Item(1).NumberFormat = NUMFMT_TEXT ' ID
        .Item(2).NumberFormat = NUMFMT_TEXT ' KeyID
        .Item(3).NumberFormat = NUMFMT_TEXT ' FieldID
        .Item(6).NumberFormat = NUMFMT_DATE ' CTime
        .Item(7).NumberFormat = NUMFMT_DATE ' MTime
        .Item(8).NumberFormat = NUMFMT_DATE ' DTime
        .Item(9).NumberFormat = NUMFMT_TEXT ' CommitID
    End With
End Sub

Public Function GetValues(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant) As Variant
    If This.Worksheet.UsedRange.Rows.Count = 1 Then Exit Sub
    
    Dim Result As Variant
    ReDim Result(1 To UBound(KeyIDs), 1 To UBound(FieldIDs))
    
    This.Worksheet.AutoFilterMode = False
    This.Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=This.Worksheet.Columns.Item(COL_LATEST_HASH)
    This.Worksheet.UsedRange.AutoFilter Field:=COL_LATEST_HASH, Criteria1:="<>"
    
    Dim FilteredRange As Range
    On Error Resume Next
    Set FilteredRange = GetWorksheetDatabodyRange(This.Worksheet).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    This.Worksheet.AutoFilterMode = False
    
    If FilteredRange Is Nothing Then Exit Function
    
    'FilteredRange.Sort Header:=xlNo, _
    '    Key1:=This.Worksheet.Columns.Item(COL_FIELD_ID), Order1:=xlAscending, _
    '    Key2:=This.Worksheet.Columns.Item(COL_KEY_ID), Order2:=xlAscending
    
    Dim FieldPartitions As Variant
    FieldPartitions = RangeHelpers.PartitionRange(FilteredRange, COL_FIELD_ID)
    
    Dim FieldIndex As Long
    For FieldIndex = 1 To UBound(FieldIDs)
        Dim FieldPartitionIndex As Long
        FieldPartitionIndex = ArraySearch.BinarySearch2(FieldPartitions, FieldIDs(FieldIndex))
        If FieldPartitionIndex > -1 Then
            Dim ThisKeyIDs As Variant
            ThisKeyIDs = FieldPartitions(FieldPartitionIndex, 4).Columns.Item(COL_KEY_ID).Value2
            
            Dim ThisValues As Variant
            ThisValues = FieldPartitions(FieldPartitionIndex, 4).Columns.Item(COL_VALUE).Value2
            
            Dim KeyIndex As Long
            For KeyIndex = 1 To UBound(KeyIDs)
                Dim SearchResult As Long
                SearchResult = ArraySearch.BinarySearch2(ThisKeyIDs, KeyIDs(KeyIndex))
                If SearchResult > -1 Then
                    Result(KeyIndex, FieldIndex) = ThisValues(SearchResult, 1)
                End If
            Next KeyIndex
        End If
    Next FieldIndex
    
    GetValues = Result
    Log.Message "GetValues() End", "RValues"
End Function

Private Function GetValue(ByVal KeyID As String, ByVal FieldID As String) As Variant
    Dim Query As String
    Query = KeyID & "x" & FieldID
        
    Dim Index As Long
    Index = ArraySearch.BinarySearch(This.LatestHash, Query)
    If Index = -1 Then Exit Function
    
    GetValue = This.Item.Item(Index - 1).Value
End Function

Public Sub UpsertValuesWithMask(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, _
    ByVal Values As Variant, ByVal ChangeMask As Variant, _
    ByVal CommitID As String)
    Dim i As Long
    For i = 1 To UBound(KeyIDs)
        Dim j As Long
        For j = 1 To UBound(FieldIDs)
            If ChangeMask(i, j) Then
                UpsertValue KeyIDs(i), FieldIDs(j), Values(i, j), CommitID
            End If
        Next j
    Next i
    
    Save
End Sub

Private Sub UpsertValue(ByVal KeyID As String, ByVal FieldID As String, ByVal Value As Variant, ByVal CommitID As String)
    Dim ExistingLatestRemoteValue As RemoteValue
    If TryGetLatestValue(KeyID, FieldID, ExistingLatestRemoteValue) Then
        If ExistingLatestRemoteValue.Value = Value Then
            Exit Sub ' No Change (A -> A)
        Else
            This.IsDirty = True ' CHG (A -> B)
            ExistingLatestRemoteValue.Latest = vbNullString
            ExistingLatestRemoteValue.ModificationTime = Application.Evaluate("=NOW()") - 0.0000012 ' 1/10th of a second
        End If
    Else
        This.IsDirty = True ' NEW (0 -> A)
    End If
    
    This.Item.Add Item:=CreateNewValue(KeyID, FieldID, Value, CommitID)
End Sub

Private Function TryGetLatestValue(ByVal KeyID As String, ByVal FieldID As String, ByRef OutRemoteValue As RemoteValue) As Boolean
    If IsEmpty(This.LatestHash) Then Exit Function
    
    Dim Query As String
    Query = KeyID & "x" & FieldID
    
    Dim Index As Long
    Index = ArraySearch.BinarySearch(This.LatestHash, Query)
    If Index < 1 Then Exit Function
    
    Set OutRemoteValue = This.Item.Item(Index - 1)
    TryGetLatestValue = True
End Function

Private Function CreateNewValue(ByVal KeyID As String, ByVal FieldID As String, ByVal Value As Variant, ByVal CommitID As String) As RemoteValue
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

' Returns an array containing RemoteValue elements for the given combination of KeyID and FieldID.
' Returns Nothing if it cannot find any RemoteValues.
Public Function GetAllValues(ByVal KeyID As String, ByVal FieldID As String) As Variant
    If This.Worksheet.UsedRange.Rows.Count = 1 Then Exit Function
    
    This.Worksheet.AutoFilterMode = False
    This.Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=This.Worksheet.Columns.Item(COL_KEY_ID), Order1:=xlAscending, _
        Key2:=This.Worksheet.Columns.Item(COL_FIELD_ID), Order2:=xlAscending, _
        Key3:=This.Worksheet.Columns.Item(COL_MTIME), Order3:=xlDescending
        
    Dim KeyPartitions As Variant
    KeyPartitions = RangeHelpers.PartitionRange(GetWorksheetDatabodyRange(This.Worksheet), COL_KEY_ID)
    
    Dim KeyIndex As Long
    KeyIndex = ArraySearch.BinarySearch2(KeyPartitions, KeyID)
    If KeyIndex = -1 Then Exit Function
    
    Dim FieldPartitions As Variant
    FieldPartitions = RangeHelpers.PartitionRange(KeyPartitions(KeyIndex, 4), COL_FIELD_ID)
    
    If IsEmpty(FieldPartitions) Then Exit Function
    
    Dim FieldIndex As Long
    FieldIndex = ArraySearch.BinarySearch2(FieldPartitions, FieldID)
    If FieldIndex = -1 Then Exit Function
    
    Dim WorksheetValues As Variant
    WorksheetValues = FieldPartitions(FieldIndex, COL_VALUE).Value2
    
    Dim Result As Variant
    ReDim Result(1 To UBound(WorksheetValues, 1))
    
    Dim i As Long
    For i = 1 To UBound(WorksheetValues, 1)
        Dim NewRemoteValue As RemoteValue
        Set NewRemoteValue = New RemoteValue
        NewRemoteValue.LoadFromArray WorksheetValues, i
        Set Result(i) = NewRemoteValue
    Next i
    
    GetAllValues = Result
End Function
