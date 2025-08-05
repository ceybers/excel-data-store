VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TableMapView 
   Caption         =   "TableMapView"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "TableMapView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TableMapView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@IgnoreModule ArgumentWithIncompatibleObjectType, HungarianNotation
'@Folder "Version4.Views"
Option Explicit

Implements IView

Private Type TState
    ViewModel As TableMapVM
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    OnCancel
End Sub

Private Sub cboLocalKey_Change()
    This.ViewModel.SelectKeyByName Me.cboLocalKey.Text
    UpdateSelectedItems
End Sub

Private Sub cboRemoteKey_Change()
    This.ViewModel.SelectKeyPathByString Me.cboRemoteKey.Value
End Sub

Private Sub imgLocalTable_Click()
    frmAbout.Show
End Sub

Private Sub lvLocalFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.SelectLocalFieldByIndex Item.Index
    UpdateSelectedItems
End Sub

Private Sub lvRemoteFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.SelectRemoteFieldByIndex Item.Index
    UpdateSelectedItems
End Sub

Private Sub lvLocalFields_DblClick()
    This.ViewModel.SelectLocalFieldByIndex Me.lvLocalFields.SelectedItem.Index
    This.ViewModel.TryAutoMapSelected
    UpdateSelectedItems
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set This.ViewModel = ViewModel
    This.IsCancelled = False
    
    UpdateControls
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub UpdateControls()
    Me.txtTableName.Value = This.ViewModel.Name
    Me.txtTableID.Value = This.ViewModel.TableID
    Me.txtMapID.Value = This.ViewModel.MapID
    
    LoadListColumnsToComboBox Me.cboLocalKey, This.ViewModel.GetListColumns
    LoadComboBoxFromCollection Me.cboRemoteKey, This.ViewModel.KeyPaths
    
    LoadLocalFieldsToListView Me.lvLocalFields, This.ViewModel.LocalFields
    LoadRemoteFieldsToListView Me.lvRemoteFields, This.ViewModel.RemoteFields
End Sub

Private Sub LoadListColumnsToComboBox(ByVal ComboBox As ComboBox, ByVal ListColumns As ListColumns)
    ComboBox.Clear
    
    Dim ListColumn As ListColumn
    For Each ListColumn In ListColumns
        ComboBox.AddItem ListColumn.Name, (ListColumn.Index - 1)
    Next ListColumn
    
    If This.ViewModel.SelectedKey <> vbNullString Then
        ComboBox.Text = This.ViewModel.SelectedKey
    End If
End Sub

Private Sub LoadComboBoxFromCollection(ByVal ComboBox As ComboBox, ByVal Collection As Collection)
    ComboBox.Clear
    If Collection.Count = 0 Then Exit Sub
    
    Dim i As Long
    For i = 1 To Collection.Count
        ComboBox.AddItem Collection.Item(i), i - 1
    Next i
    
    If This.ViewModel.SelectedKeyPath <> vbNullString Then
        ComboBox.Value = This.ViewModel.SelectedKeyPath
    End If
End Sub

Private Sub LoadLocalFieldsToListView(ByVal ListView As MSComctlLib.ListView, ByVal LocalFields As Collection)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .View = lvwReport
        .ColumnHeaders.Add Text:="Local"
        .ColumnHeaders.Add Text:="Remote"
        .ColumnHeaders.Item(1).Width = (.Width / 2) - 8
        .ColumnHeaders.Item(2).Width = (.Width / 2) - 8
    End With
    
    Dim i As Long
    For i = 1 To LocalFields.Count
        Dim LocalField As MappedFieldVM
        Set LocalField = LocalFields.Item(i)
        
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(Text:=LocalField.Name)
        ListItem.ListSubItems.Add Text:=LocalField.MappedToCaption
    Next i
    
    If ListView.ListItems.Count > 1 Then
        ListView.ListItems.Item(1).Selected = True
        This.ViewModel.SelectLocalFieldByIndex 1
    End If
    
    UpdateLocalFieldsInListView ListView, LocalFields
End Sub

Private Sub UpdateLocalFieldsInListView(ByVal ListView As MSComctlLib.ListView, ByVal LocalFields As Collection)
    Dim i As Long
    For i = 1 To ListView.ListItems.Count
        Dim LocalField As MappedFieldVM
        Set LocalField = LocalFields.Item(i)
        
        With ListView.ListItems.Item(i)
            .ListSubItems.Item(1).Text = LocalField.MappedToCaption
            If LocalField.IsKey Then
                .ListSubItems.Item(1).Text = "(Key column)"
            End If
            .ListSubItems.Item(1).ForeColor = IIf(LocalField.IsMapped, RGB(0, 0, 0), RGB(128, 128, 128))
            .ForeColor = IIf(LocalField.IsKey, RGB(128, 128, 128), RGB(0, 0, 0))
        End With
    Next i
    
End Sub

Private Sub LoadRemoteFieldsToListView(ByVal ListView As MSComctlLib.ListView, ByVal RemoteFields As RemoteFields)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .View = lvwReport
        .ColumnHeaders.Add Text:="Caption"
        .ColumnHeaders.Add Text:="Name"
        .ColumnHeaders.Add Text:="Name"
        .ColumnHeaders.Item(1).Width = 96 '(.Width / 2) - 8
        .ColumnHeaders.Item(2).Width = 72 '(.Width / 3) - 8
        .ColumnHeaders.Item(3).Width = 72 '(.Width / 3) - 8
    End With
   
    Dim i As Long
    For i = 1 To RemoteFields.Count
        LoadRemoteFieldToListView ListView, RemoteFields.Item(i)
    Next i
End Sub

Private Sub LoadRemoteFieldToListView(ByVal ListView As MSComctlLib.ListView, ByVal RemoteField As RemoteField)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=RemoteField.Caption)
    
    ListItem.ListSubItems.Add Text:=RemoteField.Path
    ListItem.ListSubItems.Add Text:=RemoteField.Name
End Sub

Private Sub UpdateSelectedItems()
    If Not This.ViewModel.SelectedLocalField Is Nothing Then
        Me.lvLocalFields.SelectedItem.ListSubItems.Item(1).Text = This.ViewModel.SelectedLocalField.MappedToCaption
    End If
    
    Dim SelectedRemoteFieldIndex As Long
    SelectedRemoteFieldIndex = This.ViewModel.SelectedRemoteFieldIndex
    If SelectedRemoteFieldIndex > 0 Then
        Me.lvRemoteFields.ListItems.Item(This.ViewModel.SelectedRemoteFieldIndex).Selected = True
    End If
    
    Dim i As Long
    For i = 1 To Me.lvLocalFields.ListItems.Count
        If This.ViewModel.LocalFields.Item(i).IsKey Then
            Me.lvLocalFields.ListItems.Item(i).ForeColor = RGB(128, 128, 128)
        Else
            Me.lvLocalFields.ListItems.Item(i).ForeColor = RGB(0, 0, 0)
        End If
    Next i
    
    If Not This.ViewModel.SelectedLocalField Is Nothing Then
        Me.lvRemoteFields.Enabled = Not This.ViewModel.SelectedLocalField.IsKey
    End If
    
    If This.ViewModel.SelectedKey <> vbNullString Then
        Me.cboLocalKey.Value = This.ViewModel.SelectedKey
    End If

    'If This.ViewModel.SelectedRemoteKey <> vbNullString Then
    '    Me.cboRemoteKey.Value = This.ViewModel.SelectedRemoteKey
    'End If
    
    'If This.ViewModel.SelectedLocalFieldIndex > 0 Then
    '    Me.lvLocalFields.ListItems.Item(This.ViewModel.SelectedLocalFieldIndex).Selected = True
    'End If
    
    'If This.ViewModel.SelectedRemoteFieldIndex > 0 Then
    '    Me.lvRemoteFields.ListItems.Item(This.ViewModel.SelectedRemoteFieldIndex).Selected = True
    'End If
    
    'If This.ViewModel.SelectedLocalFieldIndex > 0 Then
    '    Me.lvLocalFields.SelectedItem.ListSubItems.Item(1) = This.ViewModel.LocalFieldVMs.Item(This.ViewModel.SelectedLocalFieldIndex).MappedToCaption
    'End If
    
    UpdateLocalFieldsInListView Me.lvLocalFields, This.ViewModel.LocalFields
End Sub
