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
'@Folder("Version4.MapApplicator")
Option Explicit

Private Const COND_FMT_FORMULA As String = "=TRUE+N(""9dd8b78c-2b33-4313-8dcf-6c1870ee0ef4"")"
Private Const COND_FMT_INTERIOR As Long = 6750105 'RGB(153,255,102)

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
    
    Dim KeyIDs As Variant
    KeyIDs = This.MappedTable.GetKeyIDs
    
    Dim FieldIDs As Variant
    FieldIDs = This.MappedTable.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    
    Dim ExistingValues As Variant
    ExistingValues = UpdateToRemote(KeyIDs, FieldIDs)
    
    Dim ChangedValues As Variant
    ChangedValues = This.Remote.Values.GetChangeMask(KeyIDs, FieldIDs, ExistingValues)
    HighlightChanges KeyIDs, FieldIDs, ChangedValues
End Sub

Private Function UpdateToRemote(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant) As Variant
    Dim Result As Variant
    ReDim Result(1 To UBound(KeyIDs), 1 To UBound(FieldIDs))
    
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
            
            Result(k, f) = vv(KeyIndex, 1)
        Next k
    Next f
    
    UpdateToRemote = Result
End Function

Private Sub HighlightChanges(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, ByVal ChangeMask As Variant)
    Dim Changes As Collection
    Set Changes = New Collection
    
    Dim f As Long
    For f = 1 To UBound(FieldIDs)
        Dim ListColumn As ListColumn
        Set ListColumn = This.MappedTable.GetListColumnFromFieldID(FieldIDs(f))
        
        Dim k As Long
        For k = 1 To UBound(KeyIDs)
            Dim KeyIndex As Long
            KeyIndex = This.MappedTable.GetRowIndexFromKeyID(KeyIDs(k))
            
            If ChangeMask(k, f) = 1 Then
                Changes.Add ListColumn.DataBodyRange.Cells.Item(KeyIndex, 1)
            End If
        Next k
    Next f
    
    If Changes.Count = 0 Then Exit Sub

    If Changes.Count = 1 Then
        HighlightRange Changes.Item(1)
        Exit Sub
    End If
    
    Dim ChangeRange As Range
    Set ChangeRange = Changes.Item(1)

    Dim i As Long
    For i = 2 To Changes.Count
        Set ChangeRange = Application.Union(ChangeRange, Changes.Item(i))
    Next i
    
    HighlightRange ChangeRange
End Sub

Private Sub RemoveHighlights(ByVal ListObject As ListObject)
    Dim FormatCondition As FormatCondition
    Do While TryGetFormatConditionByFormula(COND_FMT_FORMULA, ListObject.Range, FormatCondition)
        FormatCondition.Delete
    Loop
End Sub

Private Function TryGetFormatConditionByFormula(ByVal Formula As String, ByVal Range As Range, ByRef OutFormatCondition As FormatCondition) As Boolean
    Dim ThisFormatCondition As FormatCondition
    For Each ThisFormatCondition In Range.FormatConditions
        If ThisFormatCondition.Formula1 = Formula Then
            Set OutFormatCondition = ThisFormatCondition
            TryGetFormatConditionByFormula = True
            Exit Function
        End If
    Next ThisFormatCondition
End Function

Private Sub HighlightRange(ByVal Range As Range)
    Dim FormatCondition  As FormatCondition
    Set FormatCondition = Range.FormatConditions.Add(Type:=xlExpression, Formula1:=COND_FMT_FORMULA)
    With FormatCondition
        .Interior.Color = COND_FMT_INTERIOR
    End With
End Sub
