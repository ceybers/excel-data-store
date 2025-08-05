VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableMapVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.ViewModels"
Option Explicit

Private Type TTableMapVM
    ListObject As ListObject
    TableMap As TableMap
    Remote As Remote
    
    MappedFields As Collection
    SelectedLocalField As MappedFieldVM
    SelectedRemoteField As RemoteField
    SelectedRemoteFieldIndex As Long
End Type
Private This As TTableMapVM

Public Property Get Name() As String
    Name = This.ListObject.Name
End Property

Public Property Get TableID() As String
    TableID = This.TableMap.TableID
End Property

Public Property Get MapID() As String
    MapID = This.TableMap.MapID
End Property

Public Property Get SelectedKey() As String
    SelectedKey = This.TableMap.KeyMap.KeyColumnName
End Property

Public Property Get SelectedKeyPath() As String
    SelectedKeyPath = This.TableMap.KeyMap.KeyPath
End Property

Public Property Get SelectedLocalField() As MappedFieldVM
    Set SelectedLocalField = This.SelectedLocalField
End Property

Public Property Get SelectedRemoteFieldIndex() As Long
    SelectedRemoteFieldIndex = This.SelectedRemoteFieldIndex
End Property

Public Property Get LocalFields() As Collection
    Set LocalFields = This.MappedFields
End Property

Public Property Get RemoteFields() As RemoteFields
    Set RemoteFields = This.Remote.Fields
End Property

Public Property Get KeyPaths() As Collection
    Set KeyPaths = This.Remote.Keys.Paths
End Property

Public Sub Load(ByVal ListObject As ListObject, ByVal TableMap As TableMap, ByVal Remote As Remote)
    Set This.ListObject = ListObject
    Set This.TableMap = TableMap
    Set This.Remote = Remote
    
    Log.Message "Load() LoadMappedFields", "TableMapVM"
    LoadMappedFields
    
    Log.Message "Load() LoadMappedFields", "TableMapVM"
    UpdateMappedFields
End Sub

Public Function GetListColumns() As ListColumns
    Set GetListColumns = This.ListObject.ListColumns
End Function

Public Sub SelectKeyByName(ByVal ListColumnName As String)
    This.TableMap.KeyMap.KeyColumnName = ListColumnName
    UpdateMappedFields
End Sub

Public Sub SelectKeyPathByString(ByVal KeyPath As String)
    This.TableMap.KeyMap.KeyPath = KeyPath
End Sub

Public Sub SelectLocalFieldByIndex(ByVal Index As Long)
    Set This.SelectedLocalField = This.MappedFields.Item(Index)
    Set This.SelectedRemoteField = This.SelectedLocalField.MappedTo
    UpdateSelectedRemoteField
End Sub

Public Sub SelectRemoteFieldByIndex(ByVal Index As Long)
    If This.SelectedLocalField.IsKey Then Exit Sub
    
    Set This.SelectedRemoteField = This.Remote.Fields.Item(Index)
    This.SelectedLocalField.MapTo This.SelectedRemoteField
    This.SelectedRemoteFieldIndex = Index
End Sub

Public Sub DEBUGPrint()
    Debug.Print "TableMap="; This.TableMap.Serialize
End Sub

Private Sub Class_Initialize()
    Set This.MappedFields = New Collection
End Sub

Private Sub LoadMappedFields()
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        Dim NewMappedFieldVM As MappedFieldVM
        Set NewMappedFieldVM = New MappedFieldVM
        NewMappedFieldVM.Load ListColumn
        
        Dim ExistingMappedField As MappedField
        If This.TableMap.FieldMap.TryGetMappedFieldByName(ListColumn.Name, ExistingMappedField) Then
            NewMappedFieldVM.LoadExisting ExistingMappedField, This.Remote
        Else
            NewMappedFieldVM.MapTo This.Remote.Fields.Item(1)
        End If
        
        This.MappedFields.Add NewMappedFieldVM
    Next ListColumn
End Sub

Private Sub UpdateMappedFields()
    Dim ThisMappedFieldVM As MappedFieldVM
    For Each ThisMappedFieldVM In This.MappedFields
        If ThisMappedFieldVM.IsKey Then
            If ThisMappedFieldVM.Name <> SelectedKey Then
                ThisMappedFieldVM.MapTo This.Remote.Fields.Item(1)
            End If
        End If
        
        ThisMappedFieldVM.IsKey = (ThisMappedFieldVM.Name = SelectedKey)
    Next ThisMappedFieldVM
End Sub

Private Sub UpdateSelectedRemoteField()
    If This.SelectedLocalField.IsKey Then
        This.SelectedRemoteFieldIndex = 1
        Exit Sub
    ElseIf This.SelectedLocalField.MappedTo Is Nothing Then
        This.SelectedRemoteFieldIndex = -1
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To This.Remote.Fields.Count
        Dim RemoteField As RemoteField
        Set RemoteField = This.Remote.Fields.Item(i)
        If This.SelectedLocalField.MappedTo.ID = RemoteField.ID Then
            This.SelectedRemoteFieldIndex = i
            Exit Sub
        End If
    Next i
End Sub

Public Sub TryAutoMapSelected()
    If This.SelectedLocalField.IsKey Then Exit Sub
    
    Dim i As Long
    For i = 1 To This.Remote.Fields.Count
        Dim RemoteField As RemoteField
        Set RemoteField = This.Remote.Fields.Item(i)
        If RemoteField.Caption = This.SelectedLocalField.Name Then
            This.SelectedLocalField.MapTo RemoteField
            Exit Sub
        End If
    Next i
End Sub

Public Sub Save()
    UpdateFieldMap
    
    Dim RemoteMapID As String
    RemoteMapID = UpdateRemoteMaps
    
    UpdateRemoteTables RemoteMapID
End Sub

Private Function UpdateRemoteMaps() As String
    Dim MapSerStr As String
    MapSerStr = This.TableMap.Serialize
    
    Dim RemoteMap As RemoteMap
    Set RemoteMap = This.Remote.Maps.AddOrGet(MapSerStr, This.ListObject.Name)
    This.Remote.Maps.Save
    
    UpdateRemoteMaps = RemoteMap.ID
End Function

Private Sub UpdateRemoteTables(ByVal MapID As String)
    Dim SerializedTable As SerializedTable
    Set SerializedTable = New SerializedTable
    Set SerializedTable.ListObject = This.ListObject
    
    Dim TblSerStr As String
    TblSerStr = SerializedTable.Serialize
    
    Dim RemoteTable As RemoteTable
    Set RemoteTable = This.Remote.Tables.AddOrGet(TblSerStr, This.ListObject.Name, MapID)
    This.Remote.Tables.Save
End Sub

Private Sub UpdateFieldMap()
    This.TableMap.FieldMap.Clear
    Dim LocalField As MappedFieldVM
    For Each LocalField In This.MappedFields
        If LocalField.IsMapped Then
            This.TableMap.FieldMap.Add LocalField.Name, LocalField.MappedTo.ID
        End If
    Next LocalField
End Sub
