VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteFieldsVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.ViewModels"
Option Explicit

Private Type TRemoteFieldsVM
    RemoteFields As RemoteFields
    SelectedItem As RemoteFieldVM
    Items As Collection
    PathItems As Collection
End Type
Private This As TRemoteFieldsVM

Public Property Get Selected() As RemoteFieldVM
    Set Selected = This.SelectedItem
End Property

Public Property Get Count() As Long
    Count = This.Items.Count
End Property

Public Property Get Item(ByVal Index As Long) As RemoteFieldVM
    Set Item = This.Items.Item(Index)
End Property

Public Property Get Paths() As Collection
    Set Paths = This.PathItems
End Property

Private Sub Class_Initialize()
    Set This.Items = New Collection
    Set This.PathItems = New Collection
End Sub

Public Sub Load(ByVal RemoteFields As RemoteFields)
    Set This.RemoteFields = RemoteFields
    
    CreateRemoteFieldVMs
    CreatePathItems
End Sub

Private Sub CreateRemoteFieldVMs()
    Debug.Assert Not This.Items Is Nothing
    Debug.Assert This.Items.Count = 0
    
    Dim SortIndex As Variant
    ReDim SortIndex(1 To This.RemoteFields.Count, 1 To 2)
    Dim i As Long
    For i = 1 To This.RemoteFields.Count
        SortIndex(i, 1) = This.RemoteFields.Item(i).Path & "\" & This.RemoteFields.Item(i).Caption
        SortIndex(i, 2) = i
    Next i
    ArraySort.QuickSort2 SortIndex
    
    For i = 1 To This.RemoteFields.Count
        Dim NewRemoteFieldVM As RemoteFieldVM
        Set NewRemoteFieldVM = New RemoteFieldVM
        NewRemoteFieldVM.Load This.RemoteFields.Item(SortIndex(i, 2))
        This.Items.Add Item:=NewRemoteFieldVM, Key:=NewRemoteFieldVM.Key
    Next i
End Sub

Private Sub CreatePathItems()
    Dim InputArray As Variant
    ReDim InputArray(1 To Count)
    
    Dim i As Long
    For i = 1 To Count
        InputArray(i) = Item(i).Path
    Next i
    
    Dim AllPaths As Variant
    AllPaths = PathHelpers.CreatePathTreeItems(InputArray)
    
    For i = 1 To UBound(AllPaths)
        This.PathItems.Add AllPaths(i)
    Next i
End Sub

Public Function TrySelectByKey(ByVal Key As String) As Boolean
    If Key = "#UNMAPPED" Then
        Set This.SelectedItem = Nothing
        Exit Function
    End If
    
    If Len(Key) <> 40 Then Exit Function
    Set This.SelectedItem = This.Items.Item(Key)
    TrySelectByKey = True
End Function

Public Sub Search(ByVal Query As String)
    Dim i As Long
    For i = 1 To Count
        If UCase$(Item(i).Name) Like UCase$(Query & "*") Then
            Log.Message "Search(" & Query & ") OK", "RFieldsVM"
            Set This.SelectedItem = Item(i)
            Exit Sub
        End If
    Next i
End Sub

Public Function TryGetByCaption(ByVal Caption As String, ByRef OutRemoteField As RemoteFieldVM) As Boolean
    Dim i As Long
    For i = 1 To Count
        If Item(i).Name = Caption Then
            Set OutRemoteField = Item(i)
            TryGetByCaption = True
            Exit Function
        End If
    Next i
End Function
