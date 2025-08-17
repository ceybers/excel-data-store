VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PullDryRunQuery"
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
    Log.Message "PullDryQuery.Run()", "DryQuery", Info_Level
    
    Log.Message "This.Remote.Reload", "DryQuery", Info_Level
    This.Remote.Reload
    
    RangeHighlighter.RemoveHighlights This.MappedTable.ListObject
    
    Dim KeyIDs As Variant
    KeyIDs = This.MappedTable.GetKeyIDs
    If IsEmpty(KeyIDs) Then Exit Sub
    
    Dim FieldIDs As Variant
    FieldIDs = This.MappedTable.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    
    Dim LocalValues As Variant
    Log.Message "Run MappedTable.GetExistingValuesFromTable", "PullQuery", Info_Level
    LocalValues = This.MappedTable.GetExistingValuesFromTable(KeyIDs, FieldIDs)
    
    Dim RemoteValues As Variant
    Log.Message "PullQuery.Run GetValues", "PullQuery", Info_Level
    RemoteValues = This.Remote.Values.GetValues(KeyIDs, FieldIDs)
    
    Dim ChangeMask As Variant
    ChangeMask = GetChangeMask(LocalValues, RemoteValues)
    
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangeMask, hcCompareOnly
End Sub
