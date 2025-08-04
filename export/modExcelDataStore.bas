Attribute VB_Name = "modExcelDataStore"
'@Folder("Version4")
Option Explicit

'@EntryPoint
Public Sub TableMapUI()
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=False)
    If MappedTable Is Nothing Then
        MsgBox MSG_MAP_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    RemoteFactory.GetRemote.Reload

    Dim ViewModel As TableMapVM
    Set ViewModel = New TableMapVM
    ViewModel.Load MappedTable.ListObject, MappedTable.TableMap, RemoteFactory.GetRemote
    
    Dim View As IView
    Set View = New TableMapView
    If View.ShowDialog(ViewModel) Then
        ViewModel.Save
        Exit Sub
    Else
        Exit Sub
    End If
End Sub

'@EntryPoint
Public Sub PullAll()
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=False)
    If MappedTable Is Nothing Then
        MsgBox MSG_PULL_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    With New PullQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
End Sub

'@EntryPoint
Public Sub PullPartial()
    ' TODO Only run if Selection covers a ListObject. Don't resolve the single and only ListObject on the selection's worksheet.
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=True)
    If MappedTable Is Nothing Then
        MsgBox MSG_PULL_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    With New PullQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
End Sub

'@EntryPoint
Public Sub PullHighlightOnly()
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=True)
    If MappedTable Is Nothing Then
        MsgBox MSG_PULL_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    With New PullDryRunQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
End Sub

'@EntryPoint
Public Sub Push()
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=False)
    If MappedTable Is Nothing Then
        MsgBox MSG_PUSH_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    With New PushQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
End Sub

'@EntryPoint
Public Sub PushPartial()
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=True)
    If MappedTable Is Nothing Then
        MsgBox MSG_PUSH_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    With New PushQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
End Sub

'@EntryPoint
Public Sub DataStoreUI()
    RemoteFactory.GetRemote.Reload
    
    Dim ViewModel As RemoteViewModel
    Set ViewModel = New RemoteViewModel
    ViewModel.Load RemoteFactory.GetRemote
    
    Dim RemoteView As IView
    Set RemoteView = New RemoteView
    If RemoteView.ShowDialog(ViewModel) Then
        Exit Sub
    Else
        Exit Sub
    End If
End Sub
