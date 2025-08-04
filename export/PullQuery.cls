VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PullQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "Version4.MapApplicator"
Option Explicit

Private Const DO_DEBUG As Boolean = False

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
    
    Dim KeyIDs As Variant
    KeyIDs = This.MappedTable.GetKeyIDs
    
    Dim FieldIDs As Variant
    FieldIDs = This.MappedTable.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    
    Dim NewValues As Variant
    NewValues = This.Remote.Values.GetValues(KeyIDs, FieldIDs)
    UpdateFromRemote KeyIDs, FieldIDs, NewValues
End Sub

Private Sub UpdateFromRemote(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, ByVal NewValues As Variant)
    Dim f As Long
    For f = 1 To UBound(FieldIDs)
        Dim ListColumn As ListColumn
        Set ListColumn = This.MappedTable.GetListColumnFromFieldID(FieldIDs(f))
        
        Dim vv As Variant
        vv = ListColumn.DataBodyRange.Value2
        
        Dim k As Long
        For k = 1 To UBound(KeyIDs)
            Dim KeyIndex As Long
            KeyIndex = This.MappedTable.GetRowIndexFromKeyID(KeyIDs(k))
            
            vv(KeyIndex, 1) = NewValues(k, f)
        Next k
        
        If DO_DEBUG Then Debug.Print "PullQuery.UpdateFromRemote writing to " & ListColumn.DataBodyRange.Address(False, False)
        ListColumn.DataBodyRange.Value2 = vv
    Next f
End Sub
