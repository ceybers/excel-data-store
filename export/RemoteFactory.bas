Attribute VB_Name = "RemoteFactory"
'@Folder("Version4.Factories")
Option Explicit

Public Function GetRemote() As Remote
    Static DataStore As DataStore
    If DataStore Is Nothing Then
        Set DataStore = New DataStore
        DataStore.Load
    End If
    
    Set GetRemote = DataStore.Remote
End Function
