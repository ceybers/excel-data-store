Attribute VB_Name = "MappedTableFactory"
'@Folder "TableMapMatcher.Factories"
Option Explicit

Public Function CreateMappedTable(ByVal PartialSelection As Boolean, ByVal Resolve As Boolean) As MappedTable
    Log.Message "CreateMappedTable()", "MapTablFct"
    Dim ListObject As ListObject
    If Not TryGetSelectedListObject(ListObject) Then
        Exit Function
    End If
    
    Dim TableMapMatches As TableMapMatches
    Set TableMapMatches = New TableMapMatches
    With TableMapMatches
        .Load
        .Evaluate ListObject
    End With
    
    Dim MappedTable As MappedTable
    Set MappedTable = TableMapMatches.GetBestMappedTable
    
    With MappedTable
        If Resolve And MappedTable.TableMap.IsMapped Then
            Log.Message " Resolve & IsMapped", "MapTablFct"
            .SelectKeys Partial:=PartialSelection
            .ResolveKeyIDs RemoteFactory.GetRemote
        
            .SelectFields Partial:=PartialSelection
        End If
    End With
    
    Set CreateMappedTable = MappedTable
End Function
