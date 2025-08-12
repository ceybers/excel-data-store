VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalFieldVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Version4.ViewModels")
Option Explicit

Private Type TMappedFieldVM
    ListColumn As ListColumn
    'MappedField As MappedField
    RemoteField As RemoteField
    IsKey As Boolean
End Type
Private This As TMappedFieldVM

Public Property Get Key() As String
    Key = This.ListColumn.Name
End Property

Public Property Get IsKey() As Boolean
    IsKey = This.IsKey
End Property

Public Property Let IsKey(ByVal RHS As Boolean)
    This.IsKey = RHS
    If RHS = True Then
        Set This.RemoteField = Nothing
    End If
End Property

Public Property Get IsMapped() As Boolean
    If This.RemoteField Is Nothing Then Exit Property
    If This.RemoteField.ID = "#UNMAPPED" Then Exit Property
    IsMapped = True
End Property

Public Property Get Name() As String
    Name = This.ListColumn.Name
End Property

Public Property Get MappedToCaption() As String
    If This.RemoteField Is Nothing Then Exit Property
    MappedToCaption = This.RemoteField.Caption
End Property

Public Property Get MappedTo() As RemoteField
    Set MappedTo = This.RemoteField
End Property

Public Sub Load(ByVal ListColumn As ListColumn)
    Set This.ListColumn = ListColumn
End Sub

Public Sub MapTo(ByVal RemoteField As RemoteField)
    Set This.RemoteField = RemoteField
End Sub

Public Sub Unmap()
    Set This.RemoteField = Nothing
End Sub
