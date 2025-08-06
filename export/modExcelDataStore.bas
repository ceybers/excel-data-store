Attribute VB_Name = "modExcelDataStore"
'@Folder("Version4")
Option Explicit

'@EntryPoint
Public Sub TableMapUI()
    Log.StartLogging
    Log.Message "TableMapMatcherUI", "TMapMatchUI"
    
    Dim VM As TableMapMatcherVM
    Set VM = New TableMapMatcherVM
    VM.Load
    If VM.IsValid = False Then
        MsgBox MSG_MAP_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    VM.GetBestMappedTable
    Log.StopLogging
        
    TableMapUIWithMappedTable VM.MappedTable
End Sub

'@EntryPoint
Public Sub TableMapMatchesUI()
    Log.StartLogging
    Log.Message "TableMapMatcherUI", "TMapMatchUI"
    
    Dim VM As TableMapMatcherVM
    Set VM = New TableMapMatcherVM
    VM.Load
    If VM.IsValid = False Then
        MsgBox MSG_MAP_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    Log.Message "Entering UserForm...", "TableMapUI", UI_Level
    Dim View As IView
    Set View = New TableMapMatcher
    If View.ShowDialog(VM) Then
        Log.Message "...exited UserForm", "TableMapUI", UI_Level
        Log.Message "ViewModel.Save", "TableMapUI"
        VM.Save
        Log.StopLogging
        
        TableMapUIWithMappedTable VM.MappedTable
        Exit Sub
    Else
        Log.Message "...exited UserForm", "TableMapUI", UI_Level
        Log.StopLogging
        Exit Sub
    End If
End Sub

Private Sub TableMapUIWithMappedTable(ByVal MappedTable As MappedTable)
    Log.StartLogging
    Log.Message "TableMapUI", "TableMapUI"
    
    Log.Message "RemoteFactory.GetRemote.Reload", "TableMapUI"
    RemoteFactory.GetRemote.Reload

    Log.Message "TableMapVM.Load LO TM GR", "TableMapUI"
    Dim ViewModel As TableMapVM
    Set ViewModel = New TableMapVM
    ViewModel.Load MappedTable.ListObject, MappedTable.TableMap, RemoteFactory.GetRemote
    
    Log.Message "Entering UserForm...", "TableMapUI", UI_Level
    Dim View As IView
    Set View = New TableMapView
    If View.ShowDialog(ViewModel) Then
        Log.Message "...exited UserForm", "TableMapUI", UI_Level
        Log.Message "ViewModel.Save", "TableMapUI"
        ViewModel.Save
        Log.StopLogging
        Exit Sub
    Else
        Log.Message "...exited UserForm", "TableMapUI", UI_Level
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
Public Sub HighlightAll()
    DoHighlight PartialSelection:=False
End Sub

'@EntryPoint
Public Sub HighlightSelection()
    DoHighlight PartialSelection:=True
End Sub

'@EntryPoint
Public Sub HighlightRemove()
    If Selection.ListObject Is Nothing Then Exit Sub
    RangeHighlighter.RemoveHighlights Selection.ListObject
End Sub

'@EntryPoint
Public Sub HighlightMappedFields()
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=True, Resolve:=True)
    If MappedTable Is Nothing Then Exit Sub
    MappedTable.HighlightMappedFields
End Sub


Private Sub DoHighlight(ByVal PartialSelection As Boolean)
    Log.StartLogging
    Log.Message "DoHighlight()"
    
    Dim MappedTable As MappedTable
    Set MappedTable = MappedTableFactory.CreateMappedTable(PartialSelection:=PartialSelection, Resolve:=True)
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
    Log.Message "Entering UserForm...", "DataStoreUI", UI_Level
    If RemoteView.ShowDialog(ViewModel) Then
        Log.Message "...exited UserForm", "DataStoreUI", UI_Level
        Exit Sub
    Else
        Log.Message "...exited UserForm", "DataStoreUI", UI_Level
        Exit Sub
    End If
    Log.StopLogging
End Sub
