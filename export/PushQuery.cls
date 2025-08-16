VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PushQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "Version4.Queries"
Option Explicit

Private Type TPullQuery
    MappedTable As MappedTable
    Remote As Remote
End Type
Private This As TPullQuery

Public Property Get MappedTable() As MappedTable
    Set MappedTable = This.MappedTable
End Property

Public Property Set MappedTable(ByVal RHS As MappedTable)
    Set This.MappedTable = RHS
End Property

Public Property Get Remote() As Remote
    Set Remote = This.Remote
End Property

Public Property Set Remote(ByVal RHS As Remote)
    Set This.Remote = RHS
End Property

Public Sub Run()
    This.Remote.Reload
    
    RangeHighlighter.RemoveHighlights This.MappedTable.ListObject
    
    Dim KeyIDs As Variant
    KeyIDs = This.MappedTable.GetKeyIDs
    
    Dim FieldIDs As Variant
    FieldIDs = This.MappedTable.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    
    Dim ExistingValues As Variant
    ExistingValues = This.MappedTable.GetExistingValuesFromTable(KeyIDs, FieldIDs)
    
    Dim ChangeCount As Long
    ChangeCount = This.Remote.Values.GetChangeCount(KeyIDs, FieldIDs, ExistingValues)
    
    Dim ChangedValues As Variant
    ChangedValues = This.Remote.Values.GetChangeMask(KeyIDs, FieldIDs, ExistingValues)
    
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangedValues, hcBeforePush
    
    If ChangeCount = 0 Then
        MsgBox MSG_PUSH_NOCHANGES, vbInformation + vbOKOnly, APP_TITLE
    ElseIf vbNo = MsgBox(Replace$(MSG_PUSH_CONFIRM, "{0}", ChangeCount), vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) Then
        Exit Sub
    End If
    
    RangeHighlighter.RemoveHighlights This.MappedTable.ListObject
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangedValues, hcAfterPush
    
    Dim Commit As RemoteCommit
    Set Commit = This.Remote.Commits.Add(This.MappedTable.TableID, This.MappedTable.MapID, This.MappedTable.Caption)
    
    Dim ValuesToPush As Variant
    ValuesToPush = UpdateToRemote(KeyIDs, FieldIDs)
    
    ArrayClean.ReplaceErrorCells ValuesToPush
    
    This.Remote.Values.UpsertValues KeyIDs, FieldIDs, ValuesToPush, Commit.CommitID
End Sub

Private Function UpdateToRemote(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant) As Variant
    Log.Message "PushQuery.UpdateToRemote creating KeyIndexes", "PushQuery", Info_Level
    Dim KeyIndexes As Variant
    ReDim KeyIndexes(1 To UBound(KeyIDs))
    Dim i As Long
    For i = 1 To UBound(KeyIDs)
        KeyIndexes(i) = This.MappedTable.GetRowIndexFromKeyID(KeyIDs(i))
    Next i
    
    Log.Message "PushQuery.UpdateToRemote Looping through Fields", "PushQuery", Info_Level
    Dim Result As Variant
    ReDim Result(1 To UBound(KeyIDs), 1 To UBound(FieldIDs))
    
    Dim f As Long
    For f = 1 To UBound(FieldIDs)
        Dim ListColumn As ListColumn
        Set ListColumn = This.MappedTable.GetListColumnFromFieldID(FieldIDs(f))
        
        Dim vv As Variant
        vv = ListColumn.DataBodyRange.Value2
        
        Dim k As Long
        For k = 1 To UBound(KeyIndexes)
            Result(k, f) = vv(KeyIndexes(k), 1)
        Next k
    Next f
    
    UpdateToRemote = Result
End Function
