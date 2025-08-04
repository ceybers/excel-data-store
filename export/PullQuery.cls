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

Private Const DO_DEBUG As Boolean = True

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
    If DO_DEBUG Then Debug.Print "PullQuery.Run Start"
    
    This.Remote.Reload
    
    Dim KeyIDs As Variant
    If DO_DEBUG Then Debug.Print "PullQuery.Run GetKeyIDs"
    KeyIDs = This.MappedTable.GetKeyIDs
    
    Dim FieldIDs As Variant
    If DO_DEBUG Then Debug.Print "PullQuery.Run GetFields"
    FieldIDs = This.MappedTable.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    
    Dim NewValues As Variant
    If DO_DEBUG Then Debug.Print "PullQuery.Run GetValues"
    NewValues = This.Remote.Values.GetValues(KeyIDs, FieldIDs)
    UpdateFromRemote KeyIDs, FieldIDs, NewValues
    
    If DO_DEBUG Then Debug.Print "PullQuery.Run End"
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
