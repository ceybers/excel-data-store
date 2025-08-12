VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableMapMatcherVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TableMapMatcher.ViewModels"
Option Explicit

Private Type TTableMapMatcherVM
    ListObject As ListObject
    Matches As TableMapMatches
    SelectedMatch As Long
    CreateNew As Boolean
    MappedTable As MappedTable
End Type
Private This As TTableMapMatcherVM

Public Property Get MappedTable() As MappedTable
    Set MappedTable = This.MappedTable
End Property

Public Property Get CreateNew() As Boolean
    CreateNew = This.CreateNew
End Property

Public Property Let CreateNew(ByVal RHS As Boolean)
    This.CreateNew = RHS
End Property

Public Property Get ListObjectName() As String
    ListObjectName = This.ListObject.Name
End Property

Public Property Get Matches() As TableMapMatches
    Set Matches = This.Matches
End Property

Public Sub Load()
    If Not TryGetSelectedListObject(This.ListObject) Then Exit Sub
    
    Set This.Matches = New TableMapMatches
    With This.Matches
        .Load
        .Evaluate This.ListObject
    End With
    
    If This.Matches.Count > 1 Then
        This.SelectedMatch = 1
    End If
End Sub

Public Property Get IsValid() As Boolean
    If This.ListObject Is Nothing Then Exit Property
    If This.Matches.Count = 0 Then Exit Property
    IsValid = True
End Property

Public Sub SelectTableMapByIndex(ByVal Index As Long)
    Debug.Assert Not This.Matches Is Nothing
    Debug.Assert Index <= This.Matches.Count
    This.SelectedMatch = Index
End Sub

Public Function GetSelectedMatch() As TableMapMatch
    If This.SelectedMatch = 0 Then Exit Function
    Set GetSelectedMatch = This.Matches.Item(This.SelectedMatch)
End Function

Public Sub Save()
    Dim MappedTable As MappedTable
    
    If This.CreateNew Then
        Set MappedTable = New MappedTable
        MappedTable.Load This.ListObject, New TableMap
    Else
        Set MappedTable = This.Matches.GetMappedTableByIndex(This.SelectedMatch)
    End If
    
    Set This.MappedTable = MappedTable
End Sub

Public Function GetBestMappedTable() As MappedTable
    Set This.MappedTable = This.Matches.GetBestMappedTable
    Set GetBestMappedTable = This.MappedTable
End Function
