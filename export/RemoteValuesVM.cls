VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteValuesVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Version4.ViewModels")
Option Explicit

Private Type TRemoteValuesVM
    Items As Collection
End Type
Private This As TRemoteValuesVM

Public Property Get Count() As Long
    Count = This.Items.Count
End Property

Public Property Get Item(ByVal Index As Long) As RemoteValueVM
    Set Item = This.Items.Item(Index)
End Property

Private Sub Class_Initialize()
    Set This.Items = New Collection
End Sub

Public Sub Load(ByVal Values As Variant, ByVal Remote As Remote)
    Dim i As Long
    For i = 1 To UBound(Values)
        Dim NewRemoteValueVM As RemoteValueVM
        Set NewRemoteValueVM = New RemoteValueVM
        NewRemoteValueVM.Load Values(i)
        NewRemoteValueVM.ResolveCommit Remote
        This.Items.Add NewRemoteValueVM
    Next i
End Sub
