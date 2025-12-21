VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ParameterCanBeByVal
'@Folder "RemoteDataStore.Models.Plural"
Option Explicit

Private Const ARRAY_WIDTH As Long = 7

Private Type TState
    Worksheet As Worksheet
    Item As Collection ' Collection<RemoteKey>
    UniquePaths As Collection
    IsDirty As Boolean
    
    CurrentKeyPath As String
    LookupKeyIDs As Variant
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Item = New Collection
    Set This.UniquePaths = New Collection
End Sub

Public Property Get Count() As Long
    Count = This.Item.Count
End Property

Public Property Get Item(ByVal Index As Long) As RemoteKey
    Set Item = This.Item.Item(Index)
End Property

Public Property Get Paths() As Collection
    Set Paths = This.UniquePaths
End Property

Public Sub Load(ByVal Worksheet As Worksheet)
    Set This.Worksheet = Worksheet
    
    LoadFromWorksheet
    GenerateUniquePaths
End Sub

Private Sub LoadFromWorksheet()
    SortWorksheet
    
    Dim DataBodyRange As Range
    Set DataBodyRange = GetWorksheetDatabodyRange(This.Worksheet)
    If DataBodyRange Is Nothing Then Exit Sub
    
    Dim WorksheetValues As Variant
    Log.Message "RemoteKeys.LoadFromWorksheet reading from " & DataBodyRange.Address(False, False), "RKeys", Info_Level
    WorksheetValues = DataBodyRange.Value2
    
    Dim i As Long
    For i = 1 To UBound(WorksheetValues, 1)
        Dim NewRemoteKey As RemoteKey
        Set NewRemoteKey = New RemoteKey
        NewRemoteKey.LoadFromArray WorksheetValues, i
        This.Item.Add Item:=NewRemoteKey
    Next i
End Sub

Private Sub GenerateUniquePaths()
    CollectionClear This.UniquePaths
    
    If Count = 0 Then Exit Sub
    
    Dim AllPaths As Variant
    ReDim AllPaths(1 To Count) As String
    
    Dim i As Long
    For i = 1 To Count
        AllPaths(i) = Item(i).Path
    Next
    
    Dim UniquePaths As Variant
    UniquePaths = ArrayUnique.Unique(AllPaths)
    
    For i = 1 To UBound(UniquePaths)
        This.UniquePaths.Add Item:=UniquePaths(i)
    Next i
End Sub

Public Sub Save()
    Debug.Assert Not This.Worksheet Is Nothing
    
    If This.IsDirty = False Then Exit Sub
    
    ClearWorksheet
    
    Dim UpdatedValues As Variant
    ReDim UpdatedValues(1 To This.Item.Count, 1 To ARRAY_WIDTH)
    
    Dim i As Long
    For i = 1 To This.Item.Count
        Dim ThisRemoteKey As RemoteKey
        Set ThisRemoteKey = This.Item.Item(i)
        ThisRemoteKey.SaveToArray UpdatedValues, i
    Next i
    
    SetNumberFormats
    
    RangeSetValueFromVariant This.Worksheet.Range("A2"), UpdatedValues
    
    SortWorksheet
End Sub

Private Sub ClearWorksheet()
    Log.Message "RemoteKeys.ClearWorksheet", "RKeys", debug_Level
    
    This.Worksheet.AutoFilterMode = False
    If This.Worksheet.UsedRange Is Nothing Then Exit Sub
    If This.Worksheet.UsedRange.Rows.Count = 1 Then Exit Sub
    This.Worksheet.UsedRange.Offset(1).Delete
End Sub

Private Sub SetNumberFormats()
    Log.Message "RemoteKeys.SetNumberFormats", "RKeys", debug_Level
    
    With This.Worksheet.Columns
        .Item(1).NumberFormat = NUMFMT_TEXT ' ID
        .Item(5).NumberFormat = NUMFMT_DATE ' CTime
        .Item(6).NumberFormat = NUMFMT_DATE ' MTime
        .Item(7).NumberFormat = NUMFMT_DATE ' DTime
    End With
End Sub

Private Sub SortWorksheet()
    ' Column 3 = Key
    This.Worksheet.UsedRange.Sort Key1:=This.Worksheet.Columns.Item(3), Header:=xlYes
End Sub

Public Function GetIDsFromQuery(ByVal KeyPath As String, ByVal Keys As Variant) As Variant
    This.CurrentKeyPath = KeyPath
    UpdateLookupKeyValues
    
    Dim ResultCount As Long
    ResultCount = UBound(Keys, 1)
    
    Dim Result As Variant
    ReDim Result(1 To ResultCount) As String
    
    Dim i As Long
    For i = 1 To ResultCount
        Result(i) = GetOrCreateKeyID(KeyPath, Keys(i))
    Next i
    
    Save
    
    This.CurrentKeyPath = vbNullString
    GetIDsFromQuery = Result
End Function

Private Sub UpdateLookupKeyValues()
    ' Sort Path A-Z; Key A-Z
    Log.Message "UpdateLookupKeyValues Sorting Path; Key...", "RKeys"
    This.Worksheet.UsedRange.Sort Header:=xlYes, _
        Key1:=This.Worksheet.Columns.Item(2), _
        Key1:=This.Worksheet.Columns.Item(3)
    Log.Message "UpdateLookupKeyValues Sorting Path; Key... done.", "RKeys"
        
    Dim Partitions As Variant
    Partitions = RangeHelpers.PartitionRange(GetWorksheetDatabodyRange(This.Worksheet), 2)
    
    Dim i As Long
    For i = 1 To UBound(Partitions, 1)
        If Partitions(i, 1) = This.CurrentKeyPath Then
            DoUpdateLookupKeyValues Partitions, i
            Exit Sub
        End If
    Next i
End Sub

Private Sub DoUpdateLookupKeyValues(ByRef Partitions As Variant, ByVal PartitionIndex As Long)
    Log.Message "DoUpdateLookupKeyValues found: " & Partitions(PartitionIndex, 2) & " to " & Partitions(PartitionIndex, 3), "RKeys"
    Dim Range As Range
    Set Range = Partitions(PartitionIndex, 4)
    
    Dim Result As Variant
    ReDim Result(1 To Range.Rows.Count, 1 To 2)
    
    Dim vvValues As Variant
    vvValues = Range.Columns.Item(3).Value2
    Dim vvID As Variant
    vvID = Range.Columns.Item(1).Value2
    
    Dim i As Long
    For i = 1 To UBound(vvValues)
        Result(i, 1) = vvValues(i, 1)
        Result(i, 2) = vvID(i, 1)
    Next i
    
    ArraySort.QuickSort2 Result
    
    This.LookupKeyIDs = Result
End Sub

Private Function GetOrCreateKeyID(ByVal KeyPath As String, ByVal Key As String) As String
    If Key = vbNullString Then
        GetOrCreateKeyID = vbNullString
        Exit Function
    End If
    
    Dim KeyIDFast As String
    KeyIDFast = TryGetKeyIDFast(Key)
    If KeyIDFast <> vbNullString Then
        GetOrCreateKeyID = KeyIDFast
        Exit Function
    End If
    
    Dim RemoteKey As RemoteKey
    Set RemoteKey = CreateNewKey(KeyPath, Key)
    This.Item.Add Item:=RemoteKey
    This.IsDirty = True
    GetOrCreateKeyID = RemoteKey.ID
End Function

Private Function TryGetKeyIDFast(ByVal Key As String) As String
    Dim Index As Long
    Index = ArraySearch.BinarySearch2(This.LookupKeyIDs, Key)
    If Index = -1 Then Exit Function
    
    TryGetKeyIDFast = This.LookupKeyIDs(Index, 2)
End Function

Private Function CreateNewKey(ByVal KeyPath As String, ByVal Key As String) As RemoteKey
    Dim NewRemoteKey As RemoteKey
    Set NewRemoteKey = New RemoteKey
    
    With NewRemoteKey
        .Key = Key
        .Path = KeyPath
        .CreationTime = Application.Evaluate("=NOW()")
        .ID = Left$(HashSHA1(.Path & "\" & .Key), HASH_LEN)
    End With
    
    Set CreateNewKey = NewRemoteKey
End Function
