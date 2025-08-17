Attribute VB_Name = "EntryPointGuards"
'@Folder "Version4"
Option Explicit

Public Function GuardSelectionSingleCell() As Boolean
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

Public Function GuardMappedTableNoListObject(ByVal MappedTable As MappedTable) As Boolean
    If Not MappedTable Is Nothing Then Exit Function
    
    MsgBox MSG_NO_TABLE, vbExclamation + vbOKOnly, APP_TITLE
    Log.StopLogging
    GuardMappedTableNoListObject = True
End Function

Public Function GuardMappedTableProtected(ByVal MappedTable As MappedTable) As Boolean
    If MappedTable.IsProtected = False Then Exit Function
    
    MsgBox MSG_IS_PROTECTED, vbExclamation + vbOKOnly, APP_TITLE
    Log.StopLogging
    GuardMappedTableProtected = True
End Function

Public Function GuardNoSelectedListObject(ByVal ListObject As ListObject) As Boolean
    If Not ListObject Is Nothing Then Exit Function
    
    MsgBox MSG_MAP_NO_TABLE, vbInformation + vbOKOnly, APP_TITLE
    Log.StopLogging
    GuardNoSelectedListObject = True
End Function
