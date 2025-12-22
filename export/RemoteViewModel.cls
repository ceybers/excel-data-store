VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "RemoteDataStore.ViewModels"
Option Explicit

Private Type TState
    Remote As Remote
    SelectedRemoteField As RemoteField
    SelectedKeyPath As String
    DoShow As Boolean
    DoClose As Boolean
End Type
Private This As TState

Public Property Get DoShow() As Boolean
    DoShow = This.DoShow
End Property

Public Property Get DoClose() As Boolean
    DoClose = This.DoClose
End Property

Public Property Get KeyPaths() As Collection
    Set KeyPaths = This.Remote.Keys.Paths
End Property

Public Property Get Keys() As RemoteKeys2
    Set Keys = This.Remote.Keys
End Property

Public Property Get Fields() As RemoteFields
    Set Fields = This.Remote.Fields
End Property

Public Property Get SelectedRemoteField() As RemoteField
    Set SelectedRemoteField = This.SelectedRemoteField
End Property

Public Property Get SelectedKeyPath() As String
    SelectedKeyPath = This.SelectedKeyPath
End Property

Public Sub Load(ByVal Remote As Remote)
    Set This.Remote = Remote
End Sub

Public Sub SelectRemoteFieldByIndex(ByVal Index As Long)
    Set This.SelectedRemoteField = This.Remote.Fields.Item(Index)
End Sub

Public Sub SelectKeyPathByString(ByVal KeyPath As String)
    This.SelectedKeyPath = KeyPath
End Sub

Public Sub AddNew()
    With This.SelectedRemoteField
        This.Remote.Fields.AddNew .Path, .Name, .Caption
        .Path = vbNullString
        .Name = vbNullString
        .Caption = vbNullString
    End With
End Sub

Public Sub ExportKeys()
    MsgBox "NYI"
End Sub

Public Sub DoRebuild()
    This.Remote.Rebuild
End Sub

Public Sub DoSave()
    This.Remote.SaveWorkbook
    This.DoClose = True
End Sub

Public Sub Show()
    This.DoShow = True
End Sub
