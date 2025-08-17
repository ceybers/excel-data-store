Attribute VB_Name = "ListColumnHelpers"
'@Folder("Helpers.ListObject")
Option Explicit

'@Description "Tries to return the ListColumn with the given name if it exists in the ListObject."
Public Function TryGetListColumn(ByVal ListColumnName As String, ByVal ListObject As ListObject, ByRef OutListColumn As ListColumn) As Boolean
Attribute TryGetListColumn.VB_Description = "Tries to return the ListColumn with the given name if it exists in the ListObject."
    If ListColumnName = vbNullString Then Exit Function
    
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        If ListColumn.Name = ListColumnName Then
            Set OutListColumn = ListColumn
            TryGetListColumn = True
            Exit Function
        End If
    Next ListColumn
End Function

'@Description "Returns True if the List Column with the given name exists in the List Object."
Public Function ListColumnExists(ByVal ListColumnName As String, ByVal ListObject As ListObject) As Boolean
Attribute ListColumnExists.VB_Description = "Returns True if the List Column with the given name exists in the List Object."
    Dim ListColumn As ListColumn
    If TryGetListColumn(ListColumnName, ListObject, ListColumn) Then
        ListColumnExists = True
        Exit Function
    End If
End Function
