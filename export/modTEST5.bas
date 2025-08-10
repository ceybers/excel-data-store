Attribute VB_Name = "modTEST5"
'@Folder "Version062"
Option Explicit

Public Sub TableMapUIWithMappedTable2(ByVal MappedTable As MappedTable)
    Dim ViewModel As TableMapVM2
    Set ViewModel = New TableMapVM2
    ViewModel.Load MappedTable, GetRemote
    
    Dim View As IView
    Set View = New TableMapView2
    If View.ShowDialog(ViewModel) Then
        Exit Sub
    Else
        Exit Sub
    End If
End Sub

