Attribute VB_Name = "CollectionHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers.Collection"
Option Explicit

'@Description "Removes all the items in a Collection."
Public Sub CollectionClear(ByVal Collection As Collection)
Attribute CollectionClear.VB_Description = "Removes all the items in a Collection."
    If Collection Is Nothing Then Exit Sub
    Do While Collection.Count > 0
        Collection.Remove Collection.Count
    Loop
End Sub

'@Description "Returns a Range which is the Union of all the Range items in a Collection. Returns Nothing if the Collection is Nothing or if the Collection is empty."
' Assumes all items in the Collection are Ranges, and that they are all Ranges on the same Worksheet.
Public Function CollectionToRangeUnion(ByVal Collection As Collection) As Range
Attribute CollectionToRangeUnion.VB_Description = "Returns a Range which is the Union of all the Range items in a Collection. Returns Nothing if the Collection is Nothing or if the Collection is empty."
    If Collection Is Nothing Then Exit Function
    If Collection.Count = 0 Then Exit Function
    If Collection.Count = 1 Then
        Set CollectionToRangeUnion = Collection.Item(1)
        Exit Function
    End If
    
    Dim Result As Range
    Set Result = Collection.Item(1)
    
    Dim i As Long
    For i = 2 To Collection.Count
        Set Result = Application.Union(Result, Collection.Item(i))
    Next i
    
    Set CollectionToRangeUnion = Result
End Function

