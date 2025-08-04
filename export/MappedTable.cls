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
    Set KeyListColumn = This.ListObject.ListColumns.Item(This.TableMap.KeyMap.KeyColumnName)
    
    Dim SelectedKeys As Range
    If Partial Then
        Set SelectedKeys = Intersect(Selection.EntireRow, KeyListColumn.DataBodyRange)
    Else
        Set SelectedKeys = KeyListColumn.DataBodyRange
    End If
    
    'TableMap.KeyMap.Clear
    
    Dim Cell As Range
    For Each Cell In SelectedKeys.Cells
        This.TableMap.KeyMap.Add Cell.Value2, Cell.Row - This.ListObject.Range.Row
    Next Cell
End Sub

Public Sub ResolveKeyIDs(ByVal Remote As Remote)
    This.TableMap.KeyMap.ResolveIDs Remote
End Sub

Public Sub SelectFields(ByVal Partial As Boolean)
    This.TableMap.FieldMap.UnselectAllFields
    
    Dim SelectedHeaders As Range
    If Partial Then
        Set SelectedHeaders = Intersect(Selection.EntireColumn, This.ListObject.HeaderRowRange)
    Else
        Set SelectedHeaders = This.ListObject.HeaderRowRange
    End If
    If SelectedHeaders Is Nothing Then Exit Sub
    
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
