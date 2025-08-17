VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyPathsVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.ViewModels"
Option Explicit

Private Type TKeyPathsVM
    Items As Collection
    SelectedItem As String
End Type
Private This As TKeyPathsVM

Public Property Get Count() As Long
    Count = This.Items.Count
End Property

Public Property Get Item(ByVal Index As Long) As String
    Item = This.Items.Item(Index)
End Property

Public Property Get Selected() As String
    Selected = This.SelectedItem
End Property

Public Property Get IsMapped() As Boolean
    IsMapped = (This.SelectedItem <> NO_KEYPATH_SET)
End Property

Private Sub Class_Initialize()
    Set This.Items = New Collection
End Sub

Public Sub Load(ByVal Remote As Remote)
    This.Items.Add NO_KEYPATH_SET
    This.SelectedItem = NO_KEYPATH_SET
    
    Dim KeyPath As Variant
    For Each KeyPath In Remote.Keys.Paths
        This.Items.Add Item:=KeyPath, Key:=KeyPath
    Next KeyPath
End Sub

Public Sub SelectByKey(ByVal Key As String)
    This.SelectedItem = Key
End Sub

Public Sub SaveTo(ByVal TableMap As TableMap)
    Debug.Assert This.SelectedItem <> NO_KEYPATH_SET
    Debug.Assert This.SelectedItem <> vbNullString
    TableMap.KeyMap.KeyPath = This.SelectedItem
End Sub

Public Sub LoadFrom(ByVal TableMap As TableMap)
    If TableMap.KeyMap.KeyPath = vbNullString Then Exit Sub
    This.SelectedItem = TableMap.KeyMap.KeyPath
End Sub
