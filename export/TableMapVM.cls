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

Private Type TTableMapVM2
    ListObject As ListObject
    TableMap As TableMap
    Remote As Remote
    
    KeyPaths As KeyPathsVM
    LocalFields As LocalFieldsVM
    RemoteFields As RemoteFieldsVM
End Type
Private This As TTableMapVM2

Public Property Get Name() As String
    Name = This.ListObject.Name
End Property

Public Property Get KeyPaths() As KeyPathsVM
    Set KeyPaths = This.KeyPaths
End Property

Public Property Get LocalFields() As LocalFieldsVM
    Set LocalFields = This.LocalFields
End Property

Public Property Get RemoteFields() As RemoteFieldsVM
    Set RemoteFields = This.RemoteFields
End Property

Public Property Get IsValid() As Boolean
    If This.LocalFields.CanReset = False Then Exit Property
    If This.KeyPaths.IsMapped = False Then Exit Property
    IsValid = True
End Property

Public Sub Load(ByVal MappedTable As MappedTable, ByVal Remote As Remote)
    Log.Message "Load()", "TMapVM2"
    
    Set This.ListObject = MappedTable.ListObject
    Set This.TableMap = MappedTable.TableMap
    Set This.Remote = Remote
    
    Set This.KeyPaths = New KeyPathsVM
    This.KeyPaths.Load This.Remote
    This.KeyPaths.LoadFrom This.TableMap
    
    Set This.LocalFields = New LocalFieldsVM
    This.LocalFields.Load This.ListObject
    This.LocalFields.LoadFrom This.TableMap, This.Remote
    
    Set This.RemoteFields = New RemoteFieldsVM
    This.RemoteFields.Load This.Remote.Fields
    
    Log.Message "TableID=" & MappedTable.TableID, "TMapVM2"
    Log.Message "MapID=" & MappedTable.MapID, "TMapVM2"
End Sub

Public Sub SelectKeyPathByKey(ByVal Key As String)
    If This.LocalFields.Selected Is Nothing Then Exit Sub
    
    This.KeyPaths.SelectByKey Key
    This.LocalFields.Selected.IsKey = This.KeyPaths.IsMapped
    This.TableMap.KeyMap.KeyColumnName = This.LocalFields.Selected.Name
    This.TableMap.KeyMap.KeyPath = Key
End Sub

Public Sub SelectLocalByKey(ByVal Key As String)
    This.LocalFields.SelectByKey Key
    If This.LocalFields.Selected.IsMapped Then
        This.RemoteFields.TrySelectByKey This.LocalFields.Selected.MappedTo.ID
    End If
End Sub

Public Sub SelectRemoteByKey(ByVal Key As String)
    If This.RemoteFields.TrySelectByKey(Key) Then
        If This.LocalFields.Selected Is Nothing Then Exit Sub
        This.LocalFields.MapSelectedTo This.RemoteFields.Selected
    Else
        If This.LocalFields.Selected Is Nothing Then Exit Sub
        This.LocalFields.Selected.Unmap
    End If
End Sub

Public Sub DoAutoMap()
    Dim i As Long
    For i = 1 To This.LocalFields.Count
        TryAutoMap This.LocalFields.Item(i)
    Next i
End Sub

Public Function TryAutoMapSelected() As Boolean
    TryAutoMapSelected = TryAutoMap(This.LocalFields.Selected)
End Function

Public Function TryAutoMap(ByVal LocalField As LocalFieldVM) As Boolean
    Dim RemoteFieldVM As RemoteFieldVM
    If This.RemoteFields.TryGetByCaption(LocalField.Name, RemoteFieldVM) Then
        LocalField.MapTo RemoteFieldVM.RemoteField
        TryAutoMap = True
    End If
End Function

Public Sub DoResetAll()
    This.LocalFields.ResetAll
End Sub

Public Sub DoSave()
    This.KeyPaths.SaveTo This.TableMap
    This.LocalFields.SaveTo This.TableMap
    This.TableMap.SetCaption This.ListObject.Name
    
    UpdateRemoteMapsAndTables
End Sub

Private Sub UpdateRemoteMapsAndTables()
    Dim RemoteMap As RemoteMap
    Set RemoteMap = This.Remote.Maps.AddOrGet(This.TableMap.Serialize, This.TableMap.Caption)
    This.Remote.Maps.Save
    
    UpdateRemoteTables RemoteMap.ID, This.TableMap.Caption
End Sub

Private Sub UpdateRemoteTables(ByVal MapID As String, ByVal MapCaption As String)
    Dim SerializedTable As SerializedTable
    Set SerializedTable = New SerializedTable
    Set SerializedTable.ListObject = This.ListObject
    
    Dim TblSerStr As String
    TblSerStr = SerializedTable.Serialize
    
    Dim WorkbookName As String
    WorkbookName = This.ListObject.Parent.Parent.Name
    
    ' TODO Why is the ViewModel setting RemoteTable.Caption?
    Dim Caption As String
    Caption = MapCaption & " in " & WorkbookName
    
    Dim RemoteTable As RemoteTable
    Set RemoteTable = This.Remote.Tables.AddOrGet(TblSerStr, Caption, MapID)
    
    This.Remote.Tables.Save
End Sub
