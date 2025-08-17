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
    Log.Message "PushQuery.Run()", "PushQuery", Info_Level
    
    Log.Message "This.Remote.Reload", "PushQuery", Info_Level
    This.Remote.Reload
    
    RangeHighlighter.RemoveHighlights This.MappedTable.ListObject
    
    Dim KeyIDs As Variant
    KeyIDs = This.MappedTable.GetKeyIDs
    If IsEmpty(KeyIDs) Then Exit Sub
    
    Dim FieldIDs As Variant
    FieldIDs = This.MappedTable.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    
    Dim LocalValues As Variant
    Log.Message "Run MappedTable.GetExistingValuesFromTable", "PushQuery", Info_Level
    LocalValues = This.MappedTable.GetExistingValuesFromTable(KeyIDs, FieldIDs)
    
    Dim RemoteValues As Variant
    Log.Message "PullQuery.Run GetValues", "PushQuery", Info_Level
    RemoteValues = This.Remote.Values.GetValues(KeyIDs, FieldIDs)
    
    Dim ChangeMask As Variant
    ChangeMask = GetChangeMask(LocalValues, RemoteValues)
    
    Dim ChangeCount As Long
    ChangeCount = CountChanges(ChangeMask)
    
    If ChangeCount = 0 Then
        MsgBox MSG_PULL_NOCHANGES, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangeMask, hcBeforePush
    
    If ChangeCount = 0 Then
        MsgBox MSG_PUSH_NOCHANGES, vbInformation + vbOKOnly, APP_TITLE
    ElseIf vbNo = MsgBox(Replace$(MSG_PUSH_CONFIRM, "{0}", ChangeCount), vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) Then
        Exit Sub
    End If
    
    RangeHighlighter.RemoveHighlights This.MappedTable.ListObject
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangeMask, hcAfterPush
    
    Dim Commit As RemoteCommit
    Set Commit = This.Remote.Commits.Add(This.MappedTable.TableID, This.MappedTable.MapID, This.MappedTable.Caption)
    
    This.Remote.Values.UpsertValuesWithMask KeyIDs, FieldIDs, LocalValues, ChangeMask, Commit.CommitID
End Sub
