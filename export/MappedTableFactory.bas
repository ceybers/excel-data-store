Attribute VB_Name = "MappedTableFactory"
'@Folder "TableMapMatcher.Factories"
Option Explicit

'@Description "Returns True if a MappedTable is found in the Remote Data Store that matches the currently Selected ListObject and sets the variable to the MappedTable object. Returns False if not."
Public Function TryCreateBestMappedTable(ByVal Remote As Remote, _
    ByRef OutMappedTable As MappedTable) As Boolean
Attribute TryCreateBestMappedTable.VB_Description = "Returns True if a MappedTable is found in the Remote Data Store that matches the currently Selected ListObject and sets the variable to the MappedTable object. Returns False if not."
    Log.Message "TryCreateBestMappedTable() begin", "MapTblFct", Info_Level
    
    Dim ListObject As ListObject
    Log.Message " TryGetSelectedListObject()", "MapTblFct", Verbose_Level
    If Not TryGetSelectedListObject(ListObject) Then
        Log.Message " Could not find a ListObject", "MapTblFct"
        Exit Function
    End If
    
    With New TableMapMatches
        Log.Message " TableMapMatches.Load", "MapTblFct", Verbose_Level
        .Load Remote
        Log.Message " TableMapMatches.Evaluate", "MapTblFct", Verbose_Level
        .Evaluate ListObject
        Log.Message " TableMapMatches.GetBestMappedTable", "MapTblFct", Verbose_Level
        Set OutMappedTable = .GetBestMappedTable
    End With
    
    TryCreateBestMappedTable = True
    Log.Message "TryCreateBestMappedTable() end", "MapTblFct", Info_Level
End Function
