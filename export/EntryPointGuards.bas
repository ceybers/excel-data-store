Attribute VB_Name = "EntryPointGuards"
'@Folder "Version4"
Option Explicit

'@Description "Returns True if the current Selection is not a single cell selected in a Worksheet."
Public Function GuardSelectionSingleCell() As Boolean
Attribute GuardSelectionSingleCell.VB_Description = "Returns True if the current Selection is not a single cell selected in a Worksheet."
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

'@Description "Returns True if MappedTable is Nothing."
Public Function GuardMappedTableNoListObject(ByVal MappedTable As MappedTable) As Boolean
Attribute GuardMappedTableNoListObject.VB_Description = "Returns True if MappedTable is Nothing."
    If Not MappedTable Is Nothing Then Exit Function
    
    MsgBox MSG_NO_TABLE, vbExclamation + vbOKOnly, APP_TITLE
    Log.StopLogging
    GuardMappedTableNoListObject = True
End Function

'@Description "Returns True if one or more of the cells in the MappedTable's ListObject is Locked and the worksheet is Protected."
Public Function GuardMappedTableProtected(ByVal MappedTable As MappedTable) As Boolean
Attribute GuardMappedTableProtected.VB_Description = "Returns True if one or more of the cells in the MappedTable's ListObject is Locked and the worksheet is Protected."
    If MappedTable.IsProtected = False Then Exit Function
    
    MsgBox MSG_IS_PROTECTED, vbExclamation + vbOKOnly, APP_TITLE
    Log.StopLogging
    GuardMappedTableProtected = True
End Function

'@Description "Returns True if the active Window is opened in Protected View mode and cannot be edited."
Public Function GuardActiveWindowProtectedView() As Boolean
Attribute GuardActiveWindowProtectedView.VB_Description = "Returns True if the active Window is opened in Protected View mode and cannot be edited."
    If Application.ActiveProtectedViewWindow Is Nothing Then Exit Function
    
    MsgBox MSG_IS_PROTECTED_VIEW, vbExclamation + vbOKOnly, APP_TITLE
    Log.StopLogging
    GuardActiveWindowProtectedView = True
End Function

'@Description "Returns True if ListObject is Nothing."
Public Function GuardNoSelectedListObject(ByVal ListObject As ListObject) As Boolean
Attribute GuardNoSelectedListObject.VB_Description = "Returns True if ListObject is Nothing."
    If Not ListObject Is Nothing Then Exit Function
    
    MsgBox MSG_MAP_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
    Log.StopLogging
    GuardNoSelectedListObject = True
End Function
