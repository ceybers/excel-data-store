Attribute VB_Name = "ListColumnHelpers"
'@Folder("Helpers.ListObject")
Option Explicit

'@Description "Returns True if a ListColumn with the specified name exists in the ListObject and sets the provided variable to the ListColumn object. Returns False if it does not."
Public Function TryGetListColumn(ByVal ListColumnName As String, ByVal ListObject As ListObject, ByRef OutListColumn As ListColumn) As Boolean
Attribute TryGetListColumn.VB_Description = "Returns True if a ListColumn with the specified name exists in the ListObject and sets the provided variable to the ListColumn object. Returns False if it does not."
    If ListColumnName = vbNullString Then Exit Function
    If ListObject Is Nothing Then Exit Function
    
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        If ListColumn.Name = ListColumnName Then
            Set OutListColumn = ListColumn
            TryGetListColumn = True
            Exit Function
        End If
    Next ListColumn
End Function

'@Description "Returns True if a ListColumn with the specified name exists in the ListObject. Returns False if it does not."
Public Function ListColumnExists(ByVal ListColumnName As String, ByVal ListObject As ListObject) As Boolean
Attribute ListColumnExists.VB_Description = "Returns True if a ListColumn with the specified name exists in the ListObject. Returns False if it does not."
    Dim ListColumn As ListColumn
    If TryGetListColumn(ListColumnName, ListObject, ListColumn) Then
        ListColumnExists = True
        Exit Function
    End If
End Function
