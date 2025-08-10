VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalFieldsVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.ViewModels"
Option Explicit

Private Type TMappedFieldsVM
    ListObject As ListObject
    Items As Collection
    SelectedItem As MappedFieldVM
End Type
Private This As TMappedFieldsVM

Public Property Get Count() As Long
    Count = This.Items.Count
End Property

Public Property Get Item(ByVal Index As Long) As MappedFieldVM
    Set Item = This.Items.Item(Index)
End Property

Public Property Get Selected() As MappedFieldVM
    Set Selected = This.SelectedItem
End Property

Public Property Get CanReset() As Boolean
    Dim i As Long
    For i = 1 To Count
        If Item(i).IsMapped = True Then
            CanReset = True
            Exit Property
        End If
    Next i
End Property

Private Sub Class_Initialize()
    Set This.Items = New Collection
End Sub

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    CreateMappedFields
End Sub

Private Sub CreateMappedFields()
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        Dim NewMappedField As MappedFieldVM
        Set NewMappedField = New MappedFieldVM
        With NewMappedField
            .Load ListColumn
        End With
        This.Items.Add Item:=NewMappedField, Key:=NewMappedField.Name
    Next ListColumn
    
    If Count > 0 Then Set This.SelectedItem = Item(1)
End Sub

Public Sub SelectByKey(ByVal Key As String)
    Set This.SelectedItem = This.Items.Item(Key)
End Sub

Public Sub MapSelectedTo(ByVal RemoteFieldVM As RemoteFieldVM)
    If This.SelectedItem Is Nothing Then Exit Sub
    
    This.SelectedItem.MapTo RemoteFieldVM.RemoteField
End Sub

Public Sub ResetAll()
    Dim i As Long
    For i = 1 To Count
        Item(i).Unmap
    Next i
End Sub

Public Sub Search(ByVal Query As String)
    Dim i As Long
    For i = 1 To Count
        If UCase$(Item(i).Name) Like UCase$(Query & "*") Then
            Log.Message "Search(" & Query & ") OK", "LFieldsVM"
            Set This.SelectedItem = Item(i)
            Exit Sub
        End If
    Next i
End Sub

Public Sub SaveTo(ByVal TableMap As TableMap)
    TableMap.FieldMap.Clear
        
    Dim i As Long
    For i = 1 To Count
        Dim ThisItem As MappedFieldVM
        Set ThisItem = Item(i)
        
        If ThisItem.IsKey Then
            TableMap.KeyMap.KeyColumnName = ThisItem.Name
        End If
        
        If ThisItem.IsMapped Then
            TableMap.FieldMap.Add ThisItem.Name, ThisItem.MappedTo.ID
        End If
    Next i
End Sub

Public Sub LoadFrom(ByVal TableMap As TableMap, ByVal Remote As Remote)
    Dim KeyColumn As MappedFieldVM
    If TryGetByName(TableMap.KeyMap.KeyColumnName, KeyColumn) Then
        KeyColumn.IsKey = True
    End If
    
    Dim i As Long
    For i = 1 To TableMap.FieldMap.Count
        LoadMappedField TableMap.FieldMap.Item(i), Remote
    Next i
End Sub

Private Sub LoadMappedField(ByVal MappedField As MappedField, ByVal Remote As Remote)
    Dim MappedFieldVM As MappedFieldVM
    If TryGetByName(MappedField.ListColumnName, MappedFieldVM) = False Then Exit Sub
    
    Dim RemoteField As RemoteField
    If Remote.Fields.TryGetRemoteFieldFromID(MappedField.FieldID, RemoteField) = False Then Exit Sub
    
    MappedFieldVM.MapTo RemoteField
End Sub

Private Function TryGetByName(ByVal FieldName As String, ByRef OutMappedFieldVM As MappedFieldVM) As Boolean
    Dim i As Long
    For i = 1 To Count
        If Item(i).Name = FieldName Then
            Set OutMappedFieldVM = Item(i)
            TryGetByName = True
        End If
    Next i
End Function
