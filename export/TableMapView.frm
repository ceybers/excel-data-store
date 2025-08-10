VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TableMapView 
   Caption         =   "Data Store Table Mapper"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000.001
   OleObjectBlob   =   "TableMapView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TableMapView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@IgnoreModule ImplicitDefaultMemberAccess, ArgumentWithIncompatibleObjectType, HungarianNotation
'@Folder "Version4.Views"
Option Explicit

Implements IView

Private Type TState
    ViewModel As TableMapVM
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmdCancel_Click()
    OnCancel
End Sub
Private Sub cmdAutoMap_Click()
    If vbYes = MsgBox(MSG_AUTOMAP, vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) Then
        This.ViewModel.DoAutoMap
        UpdateControls
    End If
End Sub

Private Sub cmdReset_Click()
    If vbYes = MsgBox(MSG_RESETALL, vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) Then
        This.ViewModel.DoResetAll
        UpdateControls
    End If
End Sub

Private Sub cmdSaveMap_Click()
    This.ViewModel.DoSave
    Me.Hide
End Sub

Private Sub cboKeyPaths_Change()
    This.ViewModel.SelectKeyPathByKey Me.cboKeyPaths.Value
    UpdateControls
End Sub

Private Sub imgTableMap_Click()
    frmAbout.Show
End Sub

Private Sub lvMappedFields_DblClick()
    If Not Me.lvMappedFields.SelectedItem Is Nothing Then
        If This.ViewModel.TryAutoMapSelected() Then
            UpdateControls
        End If
    End If
End Sub

Private Sub lvMappedFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.SelectLocalByKey Item.Key
    UpdateControls
End Sub

Private Sub tvRemoteFields_DblClick()
    If Not Me.tvRemoteFields.SelectedItem Is Nothing Then
        This.ViewModel.SelectRemoteByKey Me.tvRemoteFields.SelectedItem.Key
        UpdateControls
    End If
End Sub

Private Sub tvRemoteFields_Expand(ByVal Node As MSComctlLib.Node)
    If PathHelpers.IsNodePath(Node.Key) Then
        Node.Image = IMG_FOLDEROPEN
    End If
End Sub

Private Sub tvRemoteFields_Collapse(ByVal Node As MSComctlLib.Node)
    If PathHelpers.IsNodePath(Node.Key) Then
        Node.Image = IMG_FOLDERCLSD
    End If
End Sub

Private Sub tvRemoteFields_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = 13 Then
        If Not Me.tvRemoteFields.SelectedItem Is Nothing Then
            This.ViewModel.SelectRemoteByKey Me.tvRemoteFields.SelectedItem.Key
        UpdateControls
        End If
        Me.lvMappedFields.SetFocus
        KeyCode = 0
    End If
End Sub

Private Sub tvRemoteFields_NodeClick(ByVal Node As MSComctlLib.Node)
    This.ViewModel.SelectRemoteByKey Node.Key
    UpdateControls
End Sub

Private Sub txtLocalFieldSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        This.ViewModel.LocalFields.Search Me.txtLocalFieldSearch.Value
        UpdateControls
        Me.lvMappedFields.SetFocus
        KeyCode = 0
    End If
End Sub

Private Sub txtRemoteFieldSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        This.ViewModel.RemoteFields.Search Me.txtRemoteFieldSearch.Value
        UpdateControls
        Me.tvRemoteFields.SetFocus
        KeyCode = 0
    End If
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
    
    InitializeControls
    UpdateControls
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub InitializeControls()
    Me.txtTableName = This.ViewModel.Name
    
    LoadKeyPathsToComboBox This.ViewModel.KeyPaths, Me.cboKeyPaths
    
    InitLocalFieldsListView Me.lvMappedFields
    LoadLocalFieldsToListView This.ViewModel.LocalFields, Me.lvMappedFields
    
    InitRemoteFieldsTreeView Me.tvRemoteFields
    LoadRemotePathsToTreeView This.ViewModel.RemoteFields, Me.tvRemoteFields
    LoadRemoteFieldsToTreeView This.ViewModel.RemoteFields, Me.tvRemoteFields
    
    UpdateControls
End Sub

Private Sub LoadKeyPathsToComboBox(ByVal KeyPaths As KeyPathsVM, ByVal ComboBox As ComboBox)
    Dim i As Long
    For i = 1 To KeyPaths.Count
        ComboBox.AddItem KeyPaths.Item(i)
    Next
    ComboBox.Enabled = False
End Sub

Private Sub InitLocalFieldsListView(ByVal ListView As ListView)
    With ListView
        .BorderStyle = ccNone
        .LabelEdit = tvwManual
        .FullRowSelect = True
        .HideSelection = False
        .View = lvwReport
        Set .SmallIcons = frmPictures16.GetImageList
    End With
    With ListView.ColumnHeaders
        .Add Text:=LV_COL_LCN, Width:=(ListView.Width / 2)
        .Add Text:=LV_COL_MAP, Width:=(ListView.Width / 2) - 16
    End With
End Sub

Private Sub LoadLocalFieldsToListView(ByVal LocalFields As LocalFieldsVM, ByVal ListView As ListView)
    Dim i As Long
    For i = 1 To LocalFields.Count
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(Key:=LocalFields.Item(i).Key, _
            Text:=LocalFields.Item(i).Name)
        ListItem.ListSubItems.Add Text:=vbNullString
    Next i
End Sub

Private Sub UpdateLocalFieldsInListView(ByVal LocalFields As LocalFieldsVM, ByVal ListView As ListView)
    Dim i As Long
    For i = 1 To LocalFields.Count
        Dim LocalField As LocalFieldVM
        Set LocalField = LocalFields.Item(i)
        
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Item(i)
        
        If This.ViewModel.LocalFields.Selected.Key = LocalField.Key Then
            ListItem.Selected = True
        Else
            ListItem.Selected = False
        End If
        
        If LocalField.IsMapped Then
            ListItem.SmallIcon = IMG_MAPPED
            ListItem.ListSubItems.Item(1) = LocalField.MappedToCaption
        ElseIf LocalField.IsKey Then
            ListItem.SmallIcon = IMG_KEY
            ListItem.ListSubItems.Item(1) = LV_CAP_UNMAPPED
        Else
            ListItem.SmallIcon = IMG_BLANK
            ListItem.ListSubItems.Item(1) = vbNullString
        End If
    Next i
End Sub

Private Sub InitRemoteFieldsTreeView(ByVal TreeView As TreeView)
    Debug.Assert TreeView.Nodes.Count = 0
    
    With TreeView
        .BorderStyle = ccNone
        .LabelEdit = tvwManual
        .HideSelection = False
        .Style = tvwTreelinesPictureText
        Set .ImageList = frmPictures16.GetImageList
        .Indentation = 16
    End With
    
    TreeView.Nodes.Add Key:=NODE_KEY_UNMAPPED, Text:=TV_CAP_UNMAPPED
    TreeView.Nodes.Add Key:=NODE_KEY_ROOT, Text:=TV_CAP_ROOT, Image:=IMG_FIELDS
End Sub

Private Sub LoadRemotePathsToTreeView(ByVal RemoteFields As RemoteFieldsVM, ByVal TreeView As TreeView)
    Dim i As Long
    For i = 1 To RemoteFields.Paths.Count
        Dim p As String
        p = RemoteFields.Paths.Item(i)
        
        Dim prev As Long
        prev = InStrRev(p, "\")
        
        Dim ParentNode As Node
        If prev = 0 Then
            Set ParentNode = TreeView.Nodes.Item(NODE_KEY_ROOT)
        Else
            Set ParentNode = TreeView.Nodes.Item(PathHelpers.PrefixNode(Left$(p, prev - 1)))
        End If
        
        Dim Caption As String
        If prev = 0 Then
            Caption = p
        Else
            Caption = Mid$(p, prev + 1, Len(p) - prev)
        End If
        
        Dim Node As Node
        Set Node = TreeView.Nodes.Add(Relative:=ParentNode, Relationship:=tvwChild, _
            Key:=PathHelpers.PrefixNode(p), Text:=Caption, Image:=IMG_FOLDEROPEN)
        Node.Expanded = True
    Next i
End Sub

Private Sub LoadRemoteFieldsToTreeView(ByVal RemoteFields As RemoteFieldsVM, ByVal TreeView As TreeView)
    Dim i As Long
    For i = 1 To RemoteFields.Count
        If RemoteFields.Item(i).Path <> vbNullString Then
            Dim ParentNode As Node
            Set ParentNode = TreeView.Nodes.Item(PathHelpers.PrefixNode(RemoteFields.Item(i).Path))
            
            TreeView.Nodes.Add Relative:=ParentNode, Relationship:=tvwChild, _
                Key:=RemoteFields.Item(i).Key, Text:=RemoteFields.Item(i).Name, _
                Image:=IMG_FIELD
        End If
    Next i
    
    TreeView.Nodes.Item(NODE_KEY_ROOT).Expanded = True
End Sub

Private Sub UpdateRemoteFieldsInTreeView(ByVal RemoteFields As RemoteFieldsVM, ByVal TreeView As TreeView)
    Dim SelectedItem As RemoteFieldVM
    Set SelectedItem = RemoteFields.Selected
    If SelectedItem Is Nothing Then Exit Sub
    
    Dim i As Long
    For i = 1 To TreeView.Nodes.Count
        Dim Node As Node
        Set Node = TreeView.Nodes.Item(i)
        If Node.Key = SelectedItem.Key Then
            Node.Selected = True
        Else
            Node.Selected = False
        End If
    Next
End Sub

Private Sub UpdateKeyPathsComboBox(ByVal ComboBox As ComboBox)
    ComboBox.Value = This.ViewModel.KeyPaths.Selected
    ComboBox.Enabled = False
    If This.ViewModel.LocalFields.Selected Is Nothing Then Exit Sub
    
    If This.ViewModel.KeyPaths.IsMapped = True Then
        If This.ViewModel.LocalFields.Selected.IsKey Then ComboBox.Enabled = True
    Else
        If This.ViewModel.LocalFields.Selected.IsMapped = False Then ComboBox.Enabled = True
    End If
End Sub

Private Sub UpdateControls()
    UpdateKeyPathsComboBox Me.cboKeyPaths
    UpdateLocalFieldsInListView This.ViewModel.LocalFields, Me.lvMappedFields
    UpdateRemoteFieldsInTreeView This.ViewModel.RemoteFields, Me.tvRemoteFields
    
    Me.cmdReset.Enabled = This.ViewModel.LocalFields.CanReset
    Me.cmdSaveMap.Enabled = This.ViewModel.IsValid
End Sub


