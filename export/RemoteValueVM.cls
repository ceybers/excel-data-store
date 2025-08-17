VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteValueVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Version4.ViewModels")
Option Explicit

Private Type TRemoteValue
    RemoteValue As RemoteValue
    RemoteCommit As RemoteCommit
End Type
Private This As TRemoteValue

Public Property Get Value() As String
    Value = This.RemoteValue.Value
End Property

Public Property Get Timestamp() As String
    Timestamp = Format$(This.RemoteValue.CreationTime, FMT_TIMESTAMP)
    If This.RemoteValue.ModificationTime = 0 Then Exit Property
    Timestamp = Format$(This.RemoteValue.ModificationTime, FMT_TIMESTAMP)
End Property

Public Property Get Commit() As String
    If This.RemoteCommit Is Nothing Then Exit Property
    Commit = This.RemoteCommit.Caption
End Property

Public Sub Load(ByVal RemoteValue As RemoteValue)
    Set This.RemoteValue = RemoteValue
End Sub

Public Sub ResolveCommit(ByVal Remote As Remote)
    Remote.Commits.TryGetByID This.RemoteValue.CommitID, This.RemoteCommit
End Sub
