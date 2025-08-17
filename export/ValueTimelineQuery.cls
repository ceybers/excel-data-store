VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueTimelineQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Version4.Queries")
Option Explicit

Private Type TPullQuery
    MappedTable As MappedTable
    Remote As Remote
    Results As Variant
End Type
Private This As TPullQuery

Public Property Get Results() As Variant
    Results = This.Results
End Property

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
    Dim KeyIDs As Variant
    KeyIDs = This.MappedTable.TableMap.KeyMap.GetKeyIDs
    If IsEmpty(KeyIDs) Then Exit Sub
    If UBound(KeyIDs) <> 1 Then Exit Sub
    
    Dim FieldIDs As Variant
    FieldIDs = This.MappedTable.TableMap.FieldMap.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    If UBound(FieldIDs) <> 1 Then Exit Sub
    
    This.Results = This.Remote.Values.GetAllValues(KeyIDs(1), FieldIDs(1))
End Sub
