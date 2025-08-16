VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MappedTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.Model"
Option Explicit

Private Type TMappedTable
    ListObject As ListObject
    TableMap As TableMap
End Type
Private This As TMappedTable

Public Property Get IsProtected() As Boolean
    IsProtected = ListObjectHelpers.TestIfProtected(This.ListObject)
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Get TableMap() As TableMap
    Set TableMap = This.TableMap
End Property

Public Property Get Caption() As String
    Caption = This.ListObject.Name & " in " & This.ListObject.Parent.Parent.Name
End Property

Public Property Get TableID() As String
    TableID = This.TableMap.TableID
End Property

Public Property Get MapID() As String
    MapID = This.TableMap.MapID
End Property

Public Function GetKeyIDs() As Variant
    GetKeyIDs = This.TableMap.KeyMap.GetKeyIDs
End Function

Public Function GetFieldIDs() As Variant
    GetFieldIDs = This.TableMap.FieldMap.GetFieldIDs
End Function

Public Sub Load(ByVal ListObject As ListObject, ByVal TableMap As TableMap)
    Set This.ListObject = ListObject
    Set This.TableMap = TableMap
End Sub

Public Sub SelectKeys(ByVal Partial As Boolean)
    Dim KeyListColumn As ListColumn
    If Not TryGetListColumn(This.TableMap.KeyMap.KeyColumnName, This.ListObject, KeyListColumn) Then Exit Sub
    
    Dim SelectedKeys As Range
    If Partial Then
        Dim SelectedRange As Range
        If Not TryGetSelectionRange(SelectedRange) Then Exit Sub
        If Not TryIntersectRanges(SelectedRange.EntireRow, KeyListColumn.DataBodyRange, SelectedKeys) Then Exit Sub
    Else
        Set SelectedKeys = KeyListColumn.DataBodyRange
    End If
    
    Dim Cell As Range
    For Each Cell In SelectedKeys.Cells
        If VarType(Cell.Value2) = vbError Then
            This.TableMap.KeyMap.Add KeyValue:=vbNullString, RowIndex:=Cell.Row - This.ListObject.Range.Row
        Else
            This.TableMap.KeyMap.Add KeyValue:=Cell.Value2, RowIndex:=Cell.Row - This.ListObject.Range.Row
        End If
    Next Cell
End Sub

Public Sub ResolveKeyIDs(ByVal Remote As Remote)
    This.TableMap.KeyMap.ResolveIDs Remote
End Sub

Public Sub SelectFields(ByVal Partial As Boolean)
    This.TableMap.FieldMap.UnselectAllFields
    
    Dim SelectedHeaders As Range
    If Partial Then
        If Not TryIntersectRanges(Selection.EntireColumn, This.ListObject.HeaderRowRange, SelectedHeaders) Then Exit Sub
    Else
        Set SelectedHeaders = This.ListObject.HeaderRowRange
    End If
    
    Dim Cell As Range
    For Each Cell In SelectedHeaders.Cells
        This.TableMap.FieldMap.SelectField Cell.Value2
    Next Cell
End Sub

Public Function GetListColumnFromFieldID(ByVal FieldID As String) As ListColumn
    Dim ListColumnName As String
    ListColumnName = This.TableMap.FieldMap.GetByFieldID(FieldID).ListColumnName
    
    Set GetListColumnFromFieldID = This.ListObject.ListColumns.Item(ListColumnName)
End Function

Public Function GetRowIndexFromKeyID(ByVal KeyID As String) As Long
    GetRowIndexFromKeyID = This.TableMap.KeyMap.GetRowFromKeyID(KeyID)
End Function

Public Function GetExistingValuesFromTable(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant) As Variant
    Dim Result As Variant
    ReDim Result(1 To UBound(KeyIDs), 1 To UBound(FieldIDs))
    
    Dim f As Long
    For f = 1 To UBound(FieldIDs)
        Dim ListColumn As ListColumn
        Set ListColumn = GetListColumnFromFieldID(FieldIDs(f))
        
        Dim vv As Variant
        vv = ListColumn.DataBodyRange.Value2
        
        Dim k As Long
        For k = 1 To UBound(KeyIDs)
            If KeyIDs(k) <> vbNullString Then
                Dim KeyIndex As Long
                KeyIndex = GetRowIndexFromKeyID(KeyIDs(k))
                
                Result(k, f) = vv(KeyIndex, 1)
            End If
        Next k
    Next f
    
    ReplaceErrorCells Result
    
    GetExistingValuesFromTable = Result
End Function

Public Sub HighlightChanges(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, ByVal ChangeMask As Variant, ByVal HighlightColor As HighlightColor)
    Dim Changes As Collection
    Set Changes = New Collection
    
    Dim f As Long
    For f = 1 To UBound(FieldIDs)
        Dim ListColumn As ListColumn
        Set ListColumn = GetListColumnFromFieldID(FieldIDs(f))
        
        Dim k As Long
        For k = 1 To UBound(KeyIDs)
            If KeyIDs(k) <> vbNullString Then
                Dim KeyIndex As Long
                KeyIndex = GetRowIndexFromKeyID(KeyIDs(k))
                
                If ChangeMask(k, f) = 1 Then
                    Changes.Add ListColumn.DataBodyRange.Cells.Item(KeyIndex, 1)
                End If
            End If
        Next k
    Next f
    
    If Changes.Count = 0 Then Exit Sub

    If Changes.Count = 1 Then
        HighlightRange Changes.Item(1), HighlightColor
        Exit Sub
    End If
    
    Dim ChangeRange As Range
    Set ChangeRange = CollectionToRangeUnion(Changes)
    
    HighlightRange ChangeRange, HighlightColor
End Sub

Public Sub HighlightMappedFields()
    RangeHighlighter.RemoveHighlights This.ListObject
    
    Dim KeyListColumn As ListColumn
    If TryGetListColumn(This.TableMap.KeyMap.KeyColumnName, This.ListObject, KeyListColumn) Then
        RangeHighlighter.HighlightRange KeyListColumn.Range.Cells.Item(1, 1), hcKey
    End If
    
    Dim MappedListColumns As Collection
    Set MappedListColumns = New Collection
    
    Dim i As Long
    For i = 1 To This.TableMap.FieldMap.Count
        Dim MappedField As MappedField
        Set MappedField = This.TableMap.FieldMap.Item(i)
        Dim ListColumn As ListColumn
        If TryGetListColumn(MappedField.ListColumnName, This.ListObject, ListColumn) Then
            MappedListColumns.Add ListColumn.Range.Cells.Item(1, 1)
        End If
    Next i
    
    If MappedListColumns.Count = 0 Then Exit Sub
    
    Dim HighlightRange As Range
    Set HighlightRange = CollectionToRangeUnion(MappedListColumns)
    
    RangeHighlighter.HighlightRange HighlightRange, hcHeaders
End Sub
