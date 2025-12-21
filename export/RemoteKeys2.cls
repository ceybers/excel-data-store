VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteKeys2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ParameterCanBeByVal
'@Folder "RemoteDataStore.Models.Plural"
Option Explicit

Private Const ARRAY_WIDTH As Long = 7
Private Const PATH_COL_INDEX As Long = 2
Private Const KEY_COL_INDEX As Long = 3

Private Type TState
    Worksheet As Worksheet
    Item As Collection
    UniquePaths As Collection

    Partitions As Variant
    KeyPaths As Scripting.Dictionary
    NewRemoteKeys As Scripting.Dictionary
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Item = New Collection
    Set This.UniquePaths = New Collection
    Set This.KeyPaths = New Scripting.Dictionary
    Set This.NewRemoteKeys = New Scripting.Dictionary
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
End Sub

Private Sub LoadFromWorksheet()
    Dim DataBodyRange As Range
    Set DataBodyRange = GetWorksheetDatabodyRange(This.Worksheet)
    If DataBodyRange Is Nothing Then Exit Sub
    
    Log.Message "RemoteKeys.LoadFromWorksheet reading from " & DataBodyRange.Address(False, False), "RKeys", Info_Level
    
    This.Partitions = RangeHelpers.PartitionRange(DataBodyRange, PATH_COL_INDEX)
    ArraySort.QuickSort2 This.Partitions
    
    Dim i As Long
    For i = 1 To UBound(This.Partitions)
        This.UniquePaths.Add Item:=This.Partitions(i, 1)
    Next i
End Sub

Private Sub Save()
    Dim NewKeyCount As Long
    NewKeyCount = This.NewRemoteKeys.Count
    If NewKeyCount = 0 Then Exit Sub
    
    Dim vv As Variant
    ReDim vv(1 To NewKeyCount, 1 To ARRAY_WIDTH)
    
    Dim ThisRemoteKey As RemoteKey
    Dim i As Long
    For i = 1 To NewKeyCount
        Set ThisRemoteKey = This.NewRemoteKeys.Items(i - 1)
        ThisRemoteKey.SaveToArray vv, i
    Next i
    
    Dim NewKeysRange As Range
    Set NewKeysRange = This.Worksheet.Cells.Item(This.Worksheet.UsedRange.Rows.Count + 1, 1)
    RangeSetValueFromVariant NewKeysRange, vv
    
    Log.Message "RemoteKeys.Save saved " & NewKeyCount & " keys", "RKeys", Info_Level
End Sub

' Main entry point
'   KeyPath = "Test\Cities"
'   Keys = Variant/String(1 to 1002)
'     Keys(1) = "New York"
'     Keys(2) = "Los Angeles"
'   GetIDsFromQuery = Variant/String(1 to 1002)
'     Result(1) = "1234ABCD"
'     Result(2) = "ABCD1234"
Public Function GetIDsFromQuery(ByVal KeyPath As String, ByVal Keys As Variant) As Variant
    Dim ResultCount As Long
    ResultCount = UBound(Keys, 1)
    Log.Message "GetIDsFromQuery (count = " & ResultCount & ")", "RKeys", Info_Level
    
    Dim Result As Variant
    ReDim Result(1 To ResultCount) As String
    
    Dim RemoteKeysDict As Scripting.Dictionary
    
    If Not TryGetKeysInPath(KeyPath, RemoteKeysDict) Then
        GetIDsFromQuery = Result
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To UBound(Keys)
        If RemoteKeysDict.Exists(Keys(i)) Then
            Result(i) = RemoteKeysDict.Item(Keys(i)).ID
        Else
            Result(i) = CreateNewKey(KeyPath, Keys(i))
            RemoteKeysDict.Add Key:=Keys(i), Item:=Result(i)
        End If
    Next i
    Log.Message " Dictionary compiled", "RKeys", Info_Level
    
    Save
    
    GetIDsFromQuery = Result
    Log.Message "GetIDsFromQuery... done", "RKeys", Info_Level
End Function

Private Function TryGetKeysInPath(ByVal KeyPath As String, ByRef OutKeyDictionary As Scripting.Dictionary) As Boolean
    Log.Message " TryGetKeysInPath...", "RKeys", Info_Level
    
    If IsEmpty(This.Partitions) Then Exit Function ' There are no existing Keys or Paths in the remote store.
    If KeyPath = vbNullString Then Exit Function ' Invalid KeyPath.
    
    If Not This.KeyPaths.Exists(KeyPath) Then
        Dim PathIndex As Long
        PathIndex = ArraySearch.BinarySearch2(This.Partitions, KeyPath)
        
        If PathIndex = -1 Then
            Set OutKeyDictionary = New Scripting.Dictionary
            TryGetKeysInPath = True
            Exit Function
            ' There are no existing Keys in this Path in the remote store.
            ' Return a blank Dictionary.
            ' We will create new Keys and add them to this in-memory dictionary and the NewRemoteKeys dict to save later.
        End If
        
        LoadKeyPath KeyPath, This.Partitions(PathIndex, 4)
    End If
    
    Set OutKeyDictionary = This.KeyPaths.Item(KeyPath)
    TryGetKeysInPath = True
    Log.Message " TryGetKeysInPath... done", "RKeys", Info_Level
End Function

Private Sub LoadKeyPath(ByVal KeyPath As String, ByVal Range As Range)
    Dim vv As Variant
    vv = Range.Value2
    
    Dim Keys As Scripting.Dictionary
    Set Keys = New Scripting.Dictionary
    
    Dim i As Long
    For i = 1 To UBound(vv, 1)
        Dim NewRemoteKey As RemoteKey
        Set NewRemoteKey = New RemoteKey
        NewRemoteKey.LoadFromArray vv, i
        Set Keys.Item(NewRemoteKey.Key) = NewRemoteKey
    Next i
    
    This.KeyPaths.Add Key:=KeyPath, Item:=Keys
    
    Log.Message "  LoadKeyPath... " & UBound(vv, 1) & " items", "RKeys", Info_Level
End Sub

Private Function CreateNewKey(ByVal KeyPath As String, ByVal Key As String) As String
    Dim NewRemoteKey As RemoteKey
    Set NewRemoteKey = New RemoteKey
    
    With NewRemoteKey
        .Key = Key
        .Path = KeyPath
        .CreationTime = Application.Evaluate("=NOW()")
        .ID = Left$(HashSHA1(.Path & "\" & .Key), HASH_LEN)
    End With
    
    This.NewRemoteKeys.Add Key:=NewRemoteKey.ID, Item:=NewRemoteKey
    CreateNewKey = NewRemoteKey.ID
End Function
