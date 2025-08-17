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
    This.Remote.Reload
    
    RemoveHighlights This.MappedTable.ListObject
    
    Log.Message "This.Remote.Reload", "PullDryRunQuery", Info_Level
    This.Remote.Reload
    
    Dim KeyIDs As Variant
    KeyIDs = This.MappedTable.GetKeyIDs
    
    Dim FieldIDs As Variant
    FieldIDs = This.MappedTable.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    
    Dim ExistingValues As Variant
    ExistingValues = This.MappedTable.GetExistingValuesFromTable(KeyIDs, FieldIDs)
    
    Dim ChangedValues As Variant
    ChangedValues = This.Remote.Values.GetChangeMask(KeyIDs, FieldIDs, ExistingValues)
    
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangedValues, hcCompareOnly
End Sub
