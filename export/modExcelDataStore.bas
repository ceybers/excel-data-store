Attribute VB_Name = "modExcelDataStore"
'@IgnoreModule FunctionReturnValueDiscarded
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

'@EntryPoint
Public Sub PullAll()
    DoPull PartialSelection:=False
End Sub

'@EntryPoint
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

'@EntryPoint
Public Sub Push()
    DoPush PartialSelection:=False
End Sub

'@EntryPoint
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

'@EntryPoint
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

'@EntryPoint
Public Sub HighlightAll()
    DoHighlight PartialSelection:=False
End Sub

'@EntryPoint
Public Sub HighlightSelection()
    DoHighlight PartialSelection:=True
End Sub

Private Sub DoHighlight(ByVal PartialSelection As Boolean)
    Log.StartLogging
    Log.Message "DoHighlight()", "DoHL"
    
    Log.Message "MappedTableFactory.TryCreateBestMappedTable", "DoHL"
    Dim MappedTable As MappedTable
    MappedTableFactory.TryCreateBestMappedTable RemoteFactory.GetRemote, MappedTable
    
    If GuardMappedTableNoListObject(MappedTable) Then Exit Sub
    If GuardMappedTableProtected(MappedTable) Then Exit Sub
    
    Log.Message "MappedTable.SelectKeysAndFields", "DoHL"
    MappedTable.SelectKeysAndFields PartialSelection:=PartialSelection
    
    Log.Message "New PullDryRunQuery", "DoHL"
    With New PullDryRunQuery
        Set .MappedTable = MappedTable
        Set .Remote = RemoteFactory.GetRemote
        .Run
    End With
    
    Log.StopLogging
End Sub

'@EntryPoint
Public Sub HighlightRemove()
    Dim ListObject As ListObject
    If Not TryGetActiveSheetListObject(ListObject) Then Exit Sub
    If TestIfProtected(ListObject) Then Exit Sub
    
    RangeHighlighter.RemoveHighlights ListObject
End Sub

'@EntryPoint
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

'@EntryPoint
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
        If ViewModel.DoClose Then RemoteFactory.GetRemote.CloseWorkbook
        Exit Sub
    Else
        Log.Message "...exited UserForm", "DataStoreUI", UI_Level
        Exit Sub
    End If
    Log.StopLogging
End Sub

'@EntryPoint
Public Sub DataStoreOpen()
    Dim Remote As Remote
    Set Remote = RemoteFactory.GetRemote
    Remote.Reload
End Sub

'@EntryPoint
Public Sub DataStoreSave()
    RemoteFactory.GetRemote.SaveWorkbook
End Sub

'@EntryPoint
Public Sub DataStoreClose()
    RemoteFactory.GetRemote.SaveWorkbook
    RemoteFactory.GetRemote.CloseWorkbook
End Sub

Private Function GuardSelectionSingleCell() As Boolean
    Dim SelectedRange As Range
    If Not TryGetSelectionRange(SelectedRange) Then
        GuardSelectionSingleCell = True
        Log.StopLogging
        Exit Function
    ElseIf SelectedRange.Cells.Count <> 1 Then
        GuardSelectionSingleCell = True
        Log.StopLogging
        Exit Function
    End If
End Function

Private Function GuardMappedTableNoListObject(ByVal MappedTable As MappedTable) As Boolean
    If Not MappedTable Is Nothing Then Exit Function
    
    MsgBox MSG_NO_TABLE, vbExclamation + vbOKOnly, APP_TITLE
    Log.StopLogging
    GuardMappedTableNoListObject = True
End Function

Private Function GuardMappedTableProtected(ByVal MappedTable As MappedTable) As Boolean
    If MappedTable.IsProtected = False Then Exit Function
    
    MsgBox MSG_IS_PROTECTED, vbExclamation + vbOKOnly, APP_TITLE
    Log.StopLogging
    GuardMappedTableProtected = True
End Function
