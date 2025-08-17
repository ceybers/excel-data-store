VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoteView 
   Caption         =   "RemoteView"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265.001
   OleObjectBlob   =   "RemoteView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoteView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ArgumentWithIncompatibleObjectType, HungarianNotation
'@Folder "RemoteDataStore.Views"
Option Explicit

Implements IView

Private Type TState
    ViewModel As RemoteViewModel
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmdAddNew_Click()
    DoAddNew
End Sub

Private Sub cmdClose_Click()
    OnCancel
End Sub

Private Sub cmdKeysExport_Click()
    This.ViewModel.ExportKeys
End Sub

Private Sub cmdRebuildIDs_Click()
    This.ViewModel.DoRebuild
    Me.cmdRebuildIDs.Enabled = False
End Sub

Private Sub cmdSave_Click()
    This.ViewModel.DoSave
    Me.cmdSave.Enabled = False
End Sub

Private Sub cmdSaveClose_Click()
    This.ViewModel.DoSave
    Me.Hide
End Sub

Private Sub cmdShow_Click()
    This.ViewModel.DoShow
    Me.cmdShow.Enabled = False
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
    Me.MultiPage1.Value = 0
End Sub

Private Sub LateLoadFields()
    LoadRemoteFieldsToListView This.ViewModel.Fields, Me.lvFields
End Sub

Private Sub LateLoadKeys()
    LoadRemoteKeyPathsToTreeView This.ViewModel.KeyPaths, Me.tvKeyPaths
End Sub

Private Sub LoadRemoteKeyPathsToTreeView(ByVal KeyPaths As Collection, ByVal TreeView As TreeView)
    With TreeView
        If .Nodes.Count > 0 Then .Nodes.Remove 1
        .HideSelection = False
        .LabelEdit = tvwManual
        .LineStyle = tvwTreeLines
        .Style = tvwTreelinesPlusMinusText
        .Indentation = 16
    End With
    
    Dim RootNode As Node
    Set RootNode = TreeView.Nodes.Add(Key:="N000", Text:="Keys")
    RootNode.Expanded = True
    
    If KeyPaths.Count = 0 Then Exit Sub
    
    RootNode.Expanded = False
    
    Dim i As Long
    For i = 1 To KeyPaths.Count
        Dim KeyPath As Variant
        KeyPath = KeyPaths.Item(i)
        TreeView.Nodes.Add Relative:=RootNode, Relationship:=tvwChild, Key:=KeyPath, Text:=KeyPath
    Next i
    
    If TreeView.Nodes.Count > 1 Then
        TreeView.Nodes.Item(2).Selected = True
        This.ViewModel.SelectKeyPathByString TreeView.Nodes.Item(2).Text
        UpdateControlsSelectedKeyPath
    End If
    
    RootNode.Expanded = True
End Sub

Private Sub LoadRemoteKeysToListView(ByVal RemoteKeys As RemoteKeys, ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .View = lvwReport
    End With
    
    With ListView.ColumnHeaders
        .Clear
        .Add Text:="ID", Width:=52 '(.Width / 5)
        .Add Text:="Name", Width:=86 '(.Width - 92) / 2
        .Add Text:="Created At", Width:=92
    End With
    
    Dim i As Long
    For i = 1 To RemoteKeys.Count
        Dim RemoteKey As RemoteKey
        Set RemoteKey = RemoteKeys.Item(i)
        
        If RemoteKey.Path = This.ViewModel.SelectedKeyPath Then
            Dim ListItem As ListItem
            Set ListItem = ListView.ListItems.Add(Key:="K#" & RemoteKey.ID, Text:=RemoteKey.ID)
            ListItem.ListSubItems.Add Text:=RemoteKey.Key
            ListItem.ListSubItems.Add Text:=RemoteKey.CreationTime
        End If
    Next i
End Sub

Private Sub LoadRemoteFieldsToListView(ByVal RemoteFields As RemoteFields, ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .View = lvwReport
    End With
    
    With ListView.ColumnHeaders
        .Clear
        .Add Text:="ID", Width:=48   '(.Width / 4) - 8
        .Add Text:="Path", Width:=48   '(.Width / 4) - 8
        .Add Text:="Name", Width:=48   '(.Width / 4) - 8
        .Add Text:="Caption", Width:=48   '(.Width / 4) - 8
    End With
    
    Dim i As Long
    For i = 1 To RemoteFields.Count
        Dim RemoteField As RemoteField
        Set RemoteField = RemoteFields.Item(i)
    
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(Text:=RemoteField.ID)
        With ListItem.ListSubItems
            .Add Text:=RemoteField.Path
            .Add Text:=RemoteField.Name
            .Add Text:=RemoteField.Caption
        End With
    Next i
    
    If ListView.ListItems.Count > 0 Then
        This.ViewModel.SelectRemoteFieldByIndex 1
        UpdateControlsSelectedRemoteField
    End If
End Sub

'Private Sub lvKeys_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    ListViewColumnClickSort Me.lvKeys, ColumnHeader.Index
'End Sub

'Private Sub lvFields_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' TODO FIX Sorting breaks the Selection/Details feature
    'ListViewColumnClickSort Me.lvFields, ColumnHeader.Index
'End Sub

Private Sub lvFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.SelectRemoteFieldByIndex Item.Index
    UpdateControlsSelectedRemoteField
End Sub

Private Sub UpdateControlsSelectedRemoteField()
    If This.ViewModel.SelectedRemoteField Is Nothing Then Exit Sub
    
    With This.ViewModel.SelectedRemoteField
        Me.txtFieldID.Value = .ID
        Me.txtFieldPath.Value = .Path
        Me.txtFieldName.Value = .Name
        Me.txtFieldCaption.Value = .Caption
        Me.txtFieldCreatedAt.Value = .CreationTime
        Me.cmdAddNew.Enabled = (.ID = FIELD_ID_ADDNEW)
        If (.ID = FIELD_ID_ADDNEW) Then Me.txtFieldCaption.Value = vbNullString
        
        Me.txtFieldPath.Locked = (.ID = FIELD_ID_UNMAPPED)
        Me.txtFieldName.Locked = (.ID = FIELD_ID_UNMAPPED)
        Me.txtFieldCaption.Locked = (.ID = FIELD_ID_UNMAPPED)
    End With
    
    Me.txtFieldID.Locked = True
    Me.txtFieldCreatedAt.Locked = True
End Sub

Private Sub MultiPage1_Change()
    ' BUG ListView on a MultiPage that is not the default Page when opened will be visually positioned at 0,0
    ' the first time the user switches to that page. If the user changes to a different page and back it will
    ' appear in the correction position. Alternatively, make it invisible then visible again.
    If Me.MultiPage1.Value = 1 Then
        If Me.tvKeyPaths.Nodes.Count = 0 Then
            LateLoadKeys
        End If
        Me.tvKeyPaths.Visible = False
        Me.tvKeyPaths.Visible = True
        Me.lvKeys.Visible = False
        Me.lvKeys.Visible = True
    ElseIf Me.MultiPage1.Value = 2 Then
        If Me.lvFields.ListItems.Count = 0 Then
            Me.lvFields.Visible = False
        End If
        LateLoadFields
        Me.lvFields.Visible = True
    End If
End Sub

Private Sub tvKeyPaths_DblClick()
    If Me.tvKeyPaths.Nodes.Item(1).Expanded = False Then
        Me.tvKeyPaths.Nodes.Item(1).Expanded = True
    End If
End Sub

Private Sub tvKeyPaths_NodeClick(ByVal Node As MSComctlLib.Node)
    This.ViewModel.SelectKeyPathByString Node.Text
    UpdateControlsSelectedKeyPath
End Sub

Private Sub UpdateControlsSelectedKeyPath()
    LoadRemoteKeysToListView This.ViewModel.Keys, Me.lvKeys
End Sub

Private Sub DoAddNew()
    If This.ViewModel.SelectedRemoteField Is Nothing Then Exit Sub
    
    With This.ViewModel.SelectedRemoteField
        .Path = Me.txtFieldPath.Value
        .Name = Me.txtFieldName.Value
        .Caption = Me.txtFieldCaption.Value
    End With
    
    This.ViewModel.AddNew
    
    UpdateControls
    
    This.ViewModel.SelectRemoteFieldByIndex Me.lvFields.ListItems.Count
    Me.lvFields.ListItems.Item(Me.lvFields.ListItems.Count).Selected = True
        
    UpdateControlsSelectedRemoteField
End Sub
