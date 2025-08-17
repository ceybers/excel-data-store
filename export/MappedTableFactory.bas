Attribute VB_Name = "MappedTableFactory"
'@Folder "TableMapMatcher.Factories"
Option Explicit

Public Function CreateMappedTable(ByVal PartialSelection As Boolean, ByVal Resolve As Boolean) As MappedTable
    Log.Message "CreateMappedTable()", "MapTablFct"
    Dim ListObject As ListObject
    
    Log.Message " TryGetSelectedListObject()", "MapTablFct"
    If Not TryGetSelectedListObject(ListObject) Then
        Exit Function
    End If
    
    Dim MappedTable As MappedTable
    With New TableMapMatches
        .Load RemoteFactory.GetRemote
        .Evaluate ListObject
        Set MappedTable = .GetBestMappedTable
    End With
    
    Set CreateMappedTable = MappedTable
    
    If Resolve = False Then Exit Function
    
    If MappedTable.TableMap.IsMapped = False Then Exit Function
    
    With MappedTable
        Log.Message " SelectKeys", "MapTablFct"
        .SelectKeys Partial:=PartialSelection
        Log.Message " ResolveKeyIDs", "MapTablFct"
        .ResolveKeyIDs RemoteFactory.GetRemote
        Log.Message " SelectFields", "MapTablFct"
        .SelectFields Partial:=PartialSelection
    End With
End Function
