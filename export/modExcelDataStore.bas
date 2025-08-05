Attribute VB_Name = "modExcelDataStore"
'@Folder("Version4")
Option Explicit

'@EntryPoint
Public Sub TableMapUI()
    Log.StartLogging
    Log.Message "TableMapUI TableMapUI", "TableMapUI"
    
    Log.Message "MappedTableFactory.CreateMappedTable", "TableMapUI"
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=False, Resolve:=False)
    If MappedTable Is Nothing Then
        MsgBox MSG_MAP_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    Log.Message "RemoteFactory.GetRemote.Reload", "TableMapUI"
    RemoteFactory.GetRemote.Reload

    Log.Message "TableMapVM.Load LO TM GR", "TableMapUI"
    Dim ViewModel As TableMapVM
    Set ViewModel = New TableMapVM
    ViewModel.Load MappedTable.ListObject, MappedTable.TableMap, RemoteFactory.GetRemote
    
    Log.Message "Entering UserForm...", "TableMapUI"
    Dim View As IView
    Set View = New TableMapView
    If View.ShowDialog(ViewModel) Then
        Log.Message "...exited UserForm", "TableMapUI"
        Log.Message "ViewModel.Save", "TableMapUI"
        ViewModel.Save
        Log.StopLogging
        Exit Sub
    Else
        Log.Message "...exited UserForm", "TableMapUI"
        Log.StopLogging
        Exit Sub
    End If
End Sub

'@EntryPoint
Public Sub PullAll()
    Log.StartLogging
    Log.Message "@EntryPoint PullAll"
    
    Log.Message "MappedTableFactory.CreateMappedTable"
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=False, Resolve:=True)
    If MappedTable Is Nothing Then
        MsgBox MSG_PULL_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    Log.Message "RemoteFactory.GetRemote.Reload", "PullAll"
    RemoteFactory.GetRemote.Reload
    
    Log.Message "New PullQuery"
    With New PullQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
    
    Log.StopLogging
End Sub

'@EntryPoint
Public Sub PullPartial()
    Log.StartLogging
    Log.Message "@EntryPoint PullPartial"
    
    ' TODO Only run if Selection covers a ListObject. Don't resolve the single and only ListObject on the selection's worksheet.
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=True, Resolve:=True)
    If MappedTable Is Nothing Then
        MsgBox MSG_PULL_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    Log.Message "RemoteFactory.GetRemote.Reload", "PullPartial"
    RemoteFactory.GetRemote.Reload
    
    With New PullQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
    
    Log.StopLogging
End Sub

'@EntryPoint
Public Sub PullHighlightOnly()
    Log.StartLogging
    Log.Message "@EntryPoint PullHighlightOnly"
    
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=True, Resolve:=True)
    If MappedTable Is Nothing Then
        MsgBox MSG_PULL_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    Log.Message "RemoteFactory.GetRemote.Reload", "PullHighlightOnly"
    RemoteFactory.GetRemote.Reload
    
    With New PullDryRunQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
    
    Log.StopLogging
End Sub

'@EntryPoint
Public Sub Push()
    Log.StartLogging
    Log.Message "@EntryPoint PushAll"
    
    Log.Message "MappedTableFactory.CreateMappedTable"
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=False, Resolve:=True)
    If MappedTable Is Nothing Then
        MsgBox MSG_PUSH_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    Log.Message "RemoteFactory.GetRemote.Reload", "Push"
    RemoteFactory.GetRemote.Reload
    
    Log.Message "New PushQuery"
    With New PushQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
    
    Log.StopLogging
End Sub

'@EntryPoint
Public Sub PushPartial()
    Log.StartLogging
    Log.Message "@EntryPoint PushPartial"
    
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=True, Resolve:=True)
    If MappedTable Is Nothing Then
        MsgBox MSG_PUSH_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    Log.Message "RemoteFactory.GetRemote.Reload", "PushPartial"
    RemoteFactory.GetRemote.Reload
    
    With New PushQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
    
    Log.StopLogging
End Sub

'@EntryPoint
Public Sub DataStoreUI()
    Log.StartLogging
    Log.Message "@EntryPoint DataStoreUI", "DataStoreUI"
    
    Log.Message "RemoteFactory.GetRemote.Reload", "DataStoreUI"
    RemoteFactory.GetRemote.Reload
    
    Log.Message "RemoteFactory.GetRemote.Show", "DataStoreUI"
    RemoteFactory.GetRemote.Show
    
    Dim ViewModel As RemoteViewModel
    Set ViewModel = New RemoteViewModel
    Log.Message "ViewModel.Load RemoteFactory.GetRemote", "DataStoreUI"
    ViewModel.Load RemoteFactory.GetRemote
    
    Dim RemoteView As IView
    Set RemoteView = New RemoteView
    Log.Message "Entering UserForm...", "DataStoreUI"
    If RemoteView.ShowDialog(ViewModel) Then
        Log.Message "...exited UserForm", "DataStoreUI"
        Exit Sub
    Else
        Log.Message "...exited UserForm", "DataStoreUI"
        Exit Sub
    End If
    Log.StopLogging
End Sub
