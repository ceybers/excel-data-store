VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableMapMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TableMapMatcher.Model")
Option Explicit

Private Type TTableMapMatches
    Matches As Collection
    TableMaps As Collection
    ListObject As ListObject
End Type
Private This As TTableMapMatches

Private Sub Class_Initialize()
    Set This.Matches = New Collection
    Set This.TableMaps = New Collection
End Sub

Public Property Get Count()
    Count = This.Matches.Count
End Property

Public Property Get Item(ByVal Index As Long) As TableMapMatch
    Set Item = This.Matches.Item(Index)
End Property

Public Sub Load()
    Log.Message "Load()", "TMapMatches"
    
    Dim Tables As RemoteTables
    Set Tables = GetRemote.Tables
    
    Dim Maps As RemoteMaps
    Set Maps = GetRemote.Maps
    
    Dim i As Long
    For i = 1 To Tables.Count
        Dim RemoteTable As RemoteTable
        Set RemoteTable = Tables.Item(i)
        
        Dim RemoteMap As RemoteMap
        Maps.TryGetByID RemoteTable.MapID, RemoteMap
        
        Dim NewTableMapMatch As TableMapMatch
        Set NewTableMapMatch = New TableMapMatch
        
        NewTableMapMatch.Load RemoteTable, RemoteMap
        
        This.Matches.Add Item:=NewTableMapMatch, Key:=RemoteTable.ID
        
        Dim TableMap As TableMap
        Set TableMap = New TableMap
        With TableMap
            .Deserialize RemoteMap.SerializedString
            .TableID = RemoteTable.ID
            .MapID = RemoteMap.ID
        End With
        This.TableMaps.Add Item:=TableMap
        
        Log.Message RemoteTable.Caption & " = " & NewTableMapMatch.Score
    Next i
    
    Log.Message "Load() end", "TMapMatches"
End Sub

Public Sub Evaluate(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    Dim i As Long
    For i = 1 To Count
        Dim ThisTableMapMatch As TableMapMatch
        Set ThisTableMapMatch = This.Matches.Item(i)
        ThisTableMapMatch.Evaluate This.ListObject
    Next i
End Sub

Public Function GetMappedTableByIndex(ByVal Index As Long) As MappedTable
    Dim MappedTable As MappedTable
    Set MappedTable = New MappedTable
    With MappedTable
        .Load This.ListObject, This.TableMaps.Item(Index)
    End With
    
    Set GetMappedTableByIndex = MappedTable
End Function

Public Function GetBestMappedTable() As MappedTable
    Dim Highscore As Long
    Dim BestIndex As Long
    
    Dim i As Long
    For i = 1 To Count
        Dim TableMapMatch As TableMapMatch
        Set TableMapMatch = This.Matches.Item(i)
        If TableMapMatch.Score > Highscore Then
            Highscore = TableMapMatch.Score
            BestIndex = i
        End If
    Next i
    
    If Highscore = 0 Or BestIndex = 0 Then
        Dim MappedTable As MappedTable
        Set MappedTable = New MappedTable
        MappedTable.Load This.ListObject, New TableMap
        Set GetBestMappedTable = MappedTable
        Exit Function
    End If
    
    Set GetBestMappedTable = GetMappedTableByIndex(BestIndex)
End Function
