VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteFieldVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.ViewModels"
Option Explicit

Private Type TRemoteFieldVM
    RemoteField As RemoteField
End Type
Private This As TRemoteFieldVM

Public Property Get Key() As String
    Key = This.RemoteField.ID
End Property

Public Property Get Name() As String
    Name = This.RemoteField.Caption
End Property

Public Property Get Path() As String
    Path = This.RemoteField.Path
End Property

Public Property Get RemoteField() As RemoteField
    Set RemoteField = This.RemoteField
End Property

Public Sub Load(ByVal RemoteField As RemoteField)
    Set This.RemoteField = RemoteField
End Sub
