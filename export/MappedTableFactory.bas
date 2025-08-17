Attribute VB_Name = "MappedTableFactory"
'@Folder "TableMapMatcher.Factories"
Option Explicit

Public Function TryCreateBestMappedTable(ByVal Remote As Remote, _
    ByRef OutMappedTable As MappedTable) As Boolean
    Log.Message "CreateMappedTable()", "MapTablFct"
    
    Dim ListObject As ListObject
    Log.Message " TryGetSelectedListObject()", "MapTablFct"
    If Not TryGetSelectedListObject(ListObject) Then
        Log.Message " Could not find a ListObject", "MapTablFct"
        Exit Function
    End If
    
    With New TableMapMatches
        Log.Message " TableMapMatches.Load", "MapTablFct"
        .Load Remote
        Log.Message " TableMapMatches.Evaluate", "MapTablFct"
        .Evaluate ListObject
        Set OutMappedTable = .GetBestMappedTable
    End With
    
    TryCreateBestMappedTable = True
End Function
