Attribute VB_Name = "modDataStore"
'@Folder "RemoteDataStore"
Option Explicit

Private Sub ZZZDataStoreMVVM()
    ' Moved to modExcelDataStore
    Dim Model As DataStore
    Set Model = New DataStore
    Model.Load
    
    Dim VM As RemoteViewModel
    Set VM = New RemoteViewModel
    VM.Load Model.Remote
    
    Dim RemoteView As IView
    Set RemoteView = New RemoteView
    If RemoteView.ShowDialog(VM) Then
        Exit Sub
    Else
        Exit Sub
    End If
End Sub
