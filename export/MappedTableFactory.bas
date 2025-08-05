Attribute VB_Name = "MappedTableFactory"
'@Folder("Version4.Factories")
Option Explicit

Public Function CreateMappedTable(ByVal PartialSelection As Boolean, ByVal Resolve As Boolean) As MappedTable
    Log.Message "CreateMappedTable()", "MapTablFct"
    Dim ListObject As ListObject
    If TryGetListObject(ListObject) = False Then
        Exit Function
    End If
    
    Dim TableMap As TableMap
    Log.Message " TryGetTableMapFromRemote()", "MapTablFct"
    If TryGetTableMapFromRemote(ListObject, TableMap) = False Then
        Log.Message "  TryGetTableMapFromRemote() = False", "MapTablFct"
        Set TableMap = New TableMap
    End If
    
    Log.Message " TableMap.TableID = " & TableMap.TableID, "MapTablFct"
    Log.Message " TableMap.MapID   = " & TableMap.MapID, "MapTablFct"
    
    Dim MappedTable As MappedTable
    Set MappedTable = New MappedTable
    With MappedTable
        Log.Message " MappedTable.Load()", "MapTablFct"
        .Load ListObject, TableMap
    
        If Resolve And TableMap.IsMapped Then
            Log.Message " Resolve & IsMapped", "MapTablFct"
            .SelectKeys Partial:=PartialSelection
            .ResolveKeyIDs RemoteFactory.GetRemote
        
            .SelectFields Partial:=PartialSelection
        End If
    End With
    
    Set CreateMappedTable = MappedTable
End Function

Private Function TryGetListObject(ByRef OutListObject As ListObject) As Boolean
    If Not TypeOf Selection Is Range Then Exit Function
    If Not Selection.ListObject Is Nothing Then
        Set OutListObject = Selection.ListObject
        TryGetListObject = True
    ElseIf Selection.Parent.ListObjects.Count = 1 Then
        Set OutListObject = Selection.Parent.ListObjects.Item(1)
        TryGetListObject = True
    End If
End Function
