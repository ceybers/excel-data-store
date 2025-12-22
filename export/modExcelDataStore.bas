Attribute VB_Name = "modExcelDataStore"
'@IgnoreModule FunctionReturnValueDiscarded
'@Folder("Version4")
Option Explicit

'@EntryPoint Map > Map Table
Public Sub TableMapUI()
    Log.StartLogging
    Log.Message "TableMapMatcherUI", "TMapMatchUI"
    
    Dim ListObject As ListObject
    TryGetSelectedListObject ListObject
    If GuardNoSelectedListObject(ListObject) Then Exit Sub
    
    Dim VM As TableMapMatcherVM
    Set VM = New TableMapMatcherVM
    With VM
        .Load ListObject
        .GetBestMappedTable
    End With
    
    Log.StopLogging
        
    TableMapUIWithMappedTable VM.MappedTable
End Sub

'@EntryPoint Map > View Maps
Public Sub TableMapMatchesUI()
    Log.StartLogging
    Log.Message "TableMapMatcherUI", "TMapMatchUI"
    
    Dim ListObject As ListObject
    TryGetSelectedListObject ListObject
    If GuardNoSelectedListObject(ListObject) Then Exit Sub
    
    Dim VM As TableMapMatcherVM
    Set VM = New TableMapMatcherVM
    VM.Load ListObject
    
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

    Log.Message "TableMapVM.Load MT GR", "TableMapUI"
    Dim ViewModel As TableMapVM
    Set ViewModel = New TableMapVM
    ViewModel.Load MappedTable, RemoteFactory.GetRemote
    
    Log.Message "Entering UserForm...", "TableMapUI", UI_Level
    Dim View As IView
    Set View = New TableMapView
    If View.ShowDialog(ViewModel) Then
        Log.Message "...exited UserForm", "TableMapUI", UI_Level
        Log.StopLogging
        Exit Sub
    Else
        Log.Message "...exited UserForm", "TableMapUI", UI_Level
        Log.StopLogging
        Exit Sub
    End If
End Sub

'@EntryPoint Pull > Pull
Public Sub PullAll()
    DoPull PartialSelection:=False
End Sub

'@EntryPoint Pull > Pull Selected
Public Sub PullPartial()
    DoPull PartialSelection:=True
End Sub

Private Sub DoPull(ByVal PartialSelection As Boolean)
    Log.StartLogging
    Log.Message "@EntryPoint DoPull", "DoPull"
    
    Log.Message "MappedTableFactory.TryCreateBestMappedTable", "DoPull"
    Dim MappedTable As MappedTable
    MappedTableFactory.TryCreateBestMappedTable RemoteFactory.GetRemote, MappedTable
    
    If GuardMappedTableNoListObject(MappedTable) Then Exit Sub
    If GuardMappedTableProtected(MappedTable) Then Exit Sub
    If GuardActiveWindowProtectedView Then Exit Sub
    
    Log.Message "MappedTable.SelectKeysAndFields", "DoPull"
    MappedTable.SelectKeysAndFields PartialSelection:=PartialSelection
    
    Log.Message "New PullQuery", "DoPull"
    With New PullQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
    
    Log.StopLogging
End Sub

'@EntryPoint Push > Push
Public Sub PushAll()
    DoPush PartialSelection:=False
End Sub

'@EntryPoint Push > Push Selected
Public Sub PushPartial()
    DoPush PartialSelection:=True
End Sub

Private Sub DoPush(ByVal PartialSelection As Boolean)
    Log.StartLogging
    Log.Message "@EntryPoint DoPush", "DoPush"
    
    Log.Message "MappedTableFactory.TryCreateBestMappedTable", "DoPush"
    Dim MappedTable As MappedTable
    MappedTableFactory.TryCreateBestMappedTable RemoteFactory.GetRemote, MappedTable
        
    If GuardMappedTableNoListObject(MappedTable) Then Exit Sub
    If GuardMappedTableProtected(MappedTable) Then Exit Sub
    
    Log.Message "MappedTable.SelectKeysAndFields", "DoPush"
    MappedTable.SelectKeysAndFields PartialSelection:=PartialSelection
    
    Log.Message "New PushQuery", "DoPush"
    With New PushQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
    
    Log.StopLogging
End Sub

'@EntryPoint Compare > Highlight Changes
Public Sub HighlightAll()
    DoHighlight PartialSelection:=False
End Sub

'@EntryPoint Compare > Select Only
Public Sub HighlightSelection()
    DoHighlight PartialSelection:=True
End Sub

Private Sub DoHighlight(ByVal PartialSelection As Boolean)
    Log.StartLogging
    Log.Message "DoHighlight()", "DoHilight", Info_Level
    
    Log.Message "MappedTableFactory.TryCreateBestMappedTable", "DoHilight", Info_Level
    Dim MappedTable As MappedTable
    MappedTableFactory.TryCreateBestMappedTable RemoteFactory.GetRemote, MappedTable
    
    If GuardMappedTableNoListObject(MappedTable) Then Exit Sub
    If GuardMappedTableProtected(MappedTable) Then Exit Sub
    If GuardActiveWindowProtectedView Then Exit Sub
    
    Log.Message "MappedTable.SelectKeysAndFields", "DoHilight", Info_Level
    MappedTable.SelectKeysAndFields PartialSelection:=PartialSelection
    
    Log.Message "New PullDryRunQuery", "DoHilight", Info_Level
    With New PullDryRunQuery
        Log.Message " Set MappedTable", "DoHilight", Verbose_Level
        Set .MappedTable = MappedTable
        Log.Message " Set Remote", "DoHilight", Verbose_Level
        Set .Remote = RemoteFactory.GetRemote
        Log.Message " PullDryRunQuery.Run()...", "DoHilight", Verbose_Level
        .Run
        Log.Message " PullDryRunQuery.Run()... done", "DoHilight", Verbose_Level
    End With
    
    Log.StopLogging
End Sub

'@EntryPoint Compare > Mapped
Public Sub HighlightMappedFields()
    Log.StartLogging
    Log.Message "@EntryPoint HighlightMappedFields", "HLMapped"
    
    Log.Message "MappedTableFactory.TryCreateBestMappedTable", "HLMapped"
    Dim MappedTable As MappedTable
    MappedTableFactory.TryCreateBestMappedTable RemoteFactory.GetRemote, MappedTable
    
    If GuardMappedTableNoListObject(MappedTable) Then Exit Sub
    If GuardMappedTableProtected(MappedTable) Then Exit Sub
    
    Log.Message "MappedTable.HighlightMappedFields", "HLMapped"
    MappedTable.HighlightMappedFields
    
    Log.StopLogging
End Sub

'@EntryPoint Compare > Clear
Public Sub HighlightRemove()
    Dim ListObject As ListObject
    If Not ListObjectHelpers.TryGetSingleListObjectInActiveSheet(ListObject) Then Exit Sub
    If ListObjectHelpers.IsListObjectProtected(ListObject) Then Exit Sub
    
    RangeHighlighter.RemoveHighlights ListObject
End Sub

'@EntryPoint History > View History
Public Sub TimelineSingle()
    Log.StartLogging
    Log.Message "@EntryPoint TLineSingle", "TimeLSngl"
    
    If GuardSelectionSingleCell Then Exit Sub
    
    Log.Message "MappedTableFactory.CreateMappedTable", "TimeLSngl"
    Dim MappedTable As MappedTable
    MappedTableFactory.TryCreateBestMappedTable RemoteFactory.GetRemote, MappedTable
    
    If GuardMappedTableNoListObject(MappedTable) Then Exit Sub
    
    MappedTable.SelectKeysAndFields PartialSelection:=True
    
    Log.Message "New ValueTimelineQuery", "TimeLSngl"
    With New ValueTimelineQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
        
        Dim VM As ValueTimelineVM
        Set VM = New ValueTimelineVM
        Log.Message "New ValueTimelineVM", "TimeLSngl"
        VM.Load .Results, RemoteFactory.GetRemote
        VM.NumberFormat = Selection.NumberFormatLocal
    End With
    
    Dim View As IView
    Set View = New ValueTimelineView
    Log.Message "Entering UserForm...", "TimeLSngl"
    If View.ShowDialog(VM) Then
        Log.Message "...exited UserForm", "TimeLSngl"
        Exit Sub
    Else
        Log.Message "...exited UserForm", "TimeLSngl"
        Exit Sub
    End If

    Log.StopLogging
End Sub

'@EntryPoint Remote > Manage Remote
Public Sub DataStoreUI()
    Log.StartLogging
    Log.Message "@EntryPoint DataStoreUI", "DataStoreUI"
    
    Log.Message "RemoteFactory.GetRemote.Reload", "DataStoreUI"
    RemoteFactory.GetRemote.Reload
    
    Dim ViewModel As RemoteViewModel
    Set ViewModel = New RemoteViewModel
    Log.Message "ViewModel.Load RemoteFactory.GetRemote", "DataStoreUI"
    ViewModel.Load RemoteFactory.GetRemote
    
    Dim RemoteView As IView
    Set RemoteView = New RemoteView
    Log.Message "Entering UserForm...", "DataStoreUI", UI_Level
    If RemoteView.ShowDialog(ViewModel) Then
        Log.Message "...exited UserForm", "DataStoreUI", UI_Level
        If ViewModel.DoShow Then RemoteFactory.GetRemote.Show
        If ViewModel.DoClose Then RemoteFactory.GetRemote.CloseWorkbook
        Exit Sub
    Else
        Log.Message "...exited UserForm", "DataStoreUI", UI_Level
        Exit Sub
    End If
    Log.StopLogging
End Sub

'@EntryPoint Remote > Open
Public Sub DataStoreOpen()
    Dim Remote As Remote
    Set Remote = RemoteFactory.GetRemote
    Remote.Reload
End Sub

'@EntryPoint Remote > Save
Public Sub DataStoreSave()
    RemoteFactory.GetRemote.SaveWorkbook
End Sub

'@EntryPoint Remote > Close
Public Sub DataStoreClose()
    RemoteFactory.GetRemote.SaveWorkbook
    RemoteFactory.GetRemote.CloseWorkbook
End Sub
