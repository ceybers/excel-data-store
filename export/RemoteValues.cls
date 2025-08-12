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
Private Const HASH_LEN As Long = 40
Private Const COL_LATEST_HASH As Long = 5

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
    LoadFromWorksheet
    LoadLatestHash
End Sub

Private Sub LoadFromWorksheet()
    SortWorksheet
    
    Do While This.Item.Count > 0
        This.Item.Remove This.Item.Count
    Loop
    
    Dim DataBodyRange As Range
    Set DataBodyRange = GetWorksheetDatabodyRange(This.Worksheet)
    If DataBodyRange Is Nothing Then Exit Sub
    
    Dim LatestHashRange As Range
    Set LatestHashRange = This.Worksheet.Columns.Item(COL_LATEST_HASH).SpecialCells(xlCellTypeConstants)
    If Not LatestHashRange Is Nothing Then
        Log.Message "RemoteValues.LoadFromWorksheet taking intersect of " & DataBodyRange.Address(False, False) & " and " & LatestHashRange.Address(False, False), "RValues", Info_Level
        If Not TryIntersectRanges(DataBodyRange, LatestHashRange.EntireRow, DataBodyRange) Then Exit Sub
    End If
    
    If DataBodyRange Is Nothing Then Exit Sub
    
    Dim WorksheetValues As Variant
    Log.Message "RemoteValues.LoadFromWorksheet reading from " & DataBodyRange.Address(False, False), "RValues", Info_Level
    WorksheetValues = DataBodyRange.Value2
    
    Dim i As Long
    For i = 1 To UBound(WorksheetValues, 1)
        'If i Mod 100 = 0 Then
        '    Log.Message "RemoteValues.LoadFromWorksheet " & i & " of " & UBound(WorksheetValues), "RemoteValues", Debug_level
        'End If
        Dim NewRemoteValue As RemoteValue
        Set NewRemoteValue = New RemoteValue
        NewRemoteValue.LoadFromArray WorksheetValues, i
        This.Item.Add Item:=NewRemoteValue
    Next i
End Sub

Private Sub LoadLatestHash()
    Dim LatestHash As Variant
    LatestHash = This.Worksheet.Columns.Item(COL_LATEST_HASH).SpecialCells(xlCellTypeConstants).Value2
    
    If VarType(LatestHash) < vbArray Then
        This.LatestHash = Empty
    Else
        This.LatestHash = ArrayTransform.ColumnToArray(LatestHash)
    End If
End Sub

Private Sub SortWorksheet()
    This.Worksheet.UsedRange.Sort Key1:=This.Worksheet.Columns.Item(COL_LATEST_HASH), Header:=xlYes
End Sub

Public Sub Save()
    Debug.Assert Not This.Worksheet Is Nothing
    
    If This.IsDirty = False Then Exit Sub
    
    ClearWorksheet
    
    Dim UpdatedValues As Variant
    ReDim UpdatedValues(1 To Count, 1 To ARRAY_WIDTH)
    
    Dim i As Long
    For i = 1 To Count
        Dim ThisRemoteValue As RemoteValue
        Set ThisRemoteValue = This.Item.Item(i)
        ThisRemoteValue.SaveToArray UpdatedValues, i
    Next i
    
    SetNumberFormats
    
    'RangeSetValueFromVariant This.Worksheet.Range("A2"), UpdatedValues
    RangeSetValueFromVariant This.Worksheet.Cells(1, 1).End(xlDown).Offset(1, 0), UpdatedValues
    
    SortWorksheet
    
    RebuildCache
End Sub

Private Sub ClearWorksheet()
    This.Worksheet.AutoFilterMode = False
    If This.Worksheet.UsedRange Is Nothing Then Exit Sub
    If This.Worksheet.UsedRange.Rows.Count = 1 Then Exit Sub
    
    This.Worksheet.UsedRange.AutoFilter Field:=5, Criteria1:="<>"
    
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

Public Function GetChangeMask(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, ByVal Values As Variant) As Variant
    Dim Result As Variant
    ReDim Result(1 To UBound(KeyIDs), 1 To UBound(FieldIDs))

    If IsEmpty(This.LatestHash) Then
        SetFullChangeMask Result
        GetChangeMask = Result
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To UBound(KeyIDs)
        Dim j As Long
        For j = 1 To UBound(FieldIDs)
            Result(i, j) = IIf(Values(i, j) = GetValue(KeyIDs(i), FieldIDs(j)), 0, 1)
        Next j
    Next i
    
    GetChangeMask = Result
End Function

Public Function GetChangeCount(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, ByVal Values As Variant) As Long
    Dim Result As Long

    If IsEmpty(This.LatestHash) Then
        GetChangeCount = UBound(KeyIDs) * UBound(FieldIDs)
        Exit Function
    End If
    
    
    Dim i As Long
    For i = 1 To UBound(KeyIDs)
        Dim j As Long
        For j = 1 To UBound(FieldIDs)
            If Values(i, j) <> GetValue(KeyIDs(i), FieldIDs(j)) Then
                Result = Result + 1
            End If
        Next j
    Next i
    
    GetChangeCount = Result
End Function

Private Sub SetFullChangeMask(ByRef ValuesVariant As Variant)
    Dim i As Long
    For i = 1 To UBound(ValuesVariant, 1)
        Dim j As Long
        For j = 1 To UBound(ValuesVariant, 2)
            ValuesVariant(i, j) = 1
        Next j
    Next i
End Sub

Public Function GetValues(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant) As Variant
    Dim Result As Variant
    ReDim Result(1 To UBound(KeyIDs), 1 To UBound(FieldIDs))
    GetValues = Result
    
    If IsEmpty(This.LatestHash) Then Exit Function

    Dim i As Long
    For i = 1 To UBound(KeyIDs)
        Dim j As Long
        For j = 1 To UBound(FieldIDs)
            Result(i, j) = GetValue(KeyIDs(i), FieldIDs(j))
        Next j
    Next i
    
    GetValues = Result
End Function

Private Function GetValue(ByVal KeyID As String, ByVal FieldID As String) As Variant
    Dim Query As String
    Query = KeyID & "x" & FieldID
        
    Dim Index As Long
    Index = ArraySearch.BinarySearch(This.LatestHash, Query)
    If Index = -1 Then Exit Function
    
    GetValue = This.Item.Item(Index - 1).Value
End Function

Public Sub UpsertValues(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, ByVal Values As Variant, ByVal CommitID As String)
    Dim i As Long
    For i = 1 To UBound(KeyIDs)
        Dim j As Long
        For j = 1 To UBound(FieldIDs)
            UpsertValue KeyIDs(i), FieldIDs(j), Values(i, j), CommitID
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
            ExistingLatestRemoteValue.ModificationTime = Application.Evaluate("=NOW()")
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

Public Function GetAllValues(ByVal KeyID As String, ByVal FieldID As String) As Variant
    This.Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=This.Worksheet.Columns.Item(2), _
        Key2:=This.Worksheet.Columns.Item(3), _
        Key3:=This.Worksheet.Columns.Item(7), Order3:=xlDescending
        
    Dim KeyPartitions As Variant
    KeyPartitions = PartitionRange(GetWorksheetDatabodyRange(This.Worksheet), 2)
    
    Dim KeyIndex As Long
    KeyIndex = ArraySearch.BinarySearch2(KeyPartitions, KeyID)
    If KeyIndex < 1 Then Exit Function
    
    Dim FieldPartitions As Variant
    FieldPartitions = PartitionRange(KeyPartitions(KeyIndex, 4), 3)
    
    Dim FieldIndex As Long
    FieldIndex = ArraySearch.BinarySearch2(FieldPartitions, FieldID)
    If FieldIndex < 1 Then Exit Function
    
    Dim Values As Variant
    Values = FieldPartitions(FieldIndex, 4).Value2
    
    Dim Values2 As Variant
    ReDim Values2(1 To UBound(Values, 1))
    
    Dim i As Long
    For i = 1 To UBound(Values, 1)
        Dim NewRemoteValue As RemoteValue
        Set NewRemoteValue = New RemoteValue
        NewRemoteValue.LoadFromArray Values, i
        Set Values2(i) = NewRemoteValue
    Next i
    
    GetAllValues = Values2
End Function
