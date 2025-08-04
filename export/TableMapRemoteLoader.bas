Attribute VB_Name = "TableMapRemoteLoader"
'@Folder "Version4.Factories"
Option Explicit

Public Function TryGetTableMapFromRemote(ByVal ListObject As ListObject, ByRef OutTableMap As TableMap) As Boolean
    Dim SerializedTable As SerializedTable
    Set SerializedTable = New SerializedTable
    Set SerializedTable.ListObject = ListObject
    
    Dim RemoteTables As RemoteTables
    Set RemoteTables = GetRemote.Tables
    
    Dim RemoteMap As RemoteMap
    
    Dim i As Long
    For i = 1 To RemoteTables.Count
        Dim RemoteTable As RemoteTable
        Set RemoteTable = RemoteTables.Item(i)
        
        SerializedTable.Deserialize RemoteTable.SerializedString
        
        If SerializedTable.TryMatch = LISTOBJECT_NAME Or SerializedTable.TryMatch = WORKSHEET_NAME Then
            If GetRemote.Maps.TryGetByID(RemoteTable.MapID, RemoteMap) = True Then
                Dim TableMap As TableMap
                If TryFitTableMap(ListObject, RemoteMap, TableMap) Then
                    TableMap.Deserialize RemoteMap.SerializedString
                    TableMap.TableID = RemoteTable.ID
                    TableMap.MapID = RemoteMap.ID
                
                    Set OutTableMap = TableMap
                    TryGetTableMapFromRemote = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function TryFitTableMap(ByVal ListObject As ListObject, ByVal RemoteMap As RemoteMap, ByRef OutTableMap As TableMap) As Boolean
    Dim TableMap As TableMap
    Set TableMap = New TableMap
    
    TableMap.Deserialize RemoteMap.SerializedString
    
    Dim KeyListColumn As ListColumn
    If TryGetListColumn(TableMap.KeyMap.KeyColumnName, ListObject, KeyListColumn) = False Then Exit Function
    
    Dim i As Long
    For i = 1 To TableMap.FieldMap.Count
        Dim FieldListColumn As ListColumn
        If TryGetListColumn(TableMap.FieldMap.Item(1).ListColumnName, ListObject, FieldListColumn) Then
            Set OutTableMap = TableMap
            TryFitTableMap = True
            Exit Function
        End If
    Next i
End Function

Private Function GetRemote() As Remote
    Dim DataStore2 As DataStore
    Set DataStore2 = New DataStore
    DataStore2.Load
    
    Set GetRemote = DataStore2.Remote
End Function
