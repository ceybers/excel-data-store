Attribute VB_Name = "CollectionHelpers"
'@Folder("Helpers")
Option Explicit

Public Sub CollectionClear(ByVal Collection As Collection)
    If Collection Is Nothing Then Exit Sub
    Do While Collection.Count > 0
        Collection.Remove Collection.Count
    Loop
End Sub

'@Description "Returns a Range which is the Union of all the Range items in a Collection."
' Assumes all items in the Collection are Ranges, and that they are all Ranges on the same Worksheet.
Public Function CollectionToRangeUnion(ByVal Collection As Collection) As Range
Attribute CollectionToRangeUnion.VB_Description = "Returns a Range which is the Union of all the Range items in a Collection."
    If Collection Is Nothing Then Exit Sub
    If Collection.Count = 0 Then Exit Sub
    
    Dim Result As Range
    Set Result = Collection.Item(1)
    
    If Collection.Count = 1 Then
        Set CollectionToRangeUnion = Result
        Exit Function
    End If
    
    Dim i As Long
    For i = 2 To Collection.Count
        Set Result = Application.Union(Result, Collection.Item(i))
    Next i
    
    Set CollectionToRangeUnion = Result
End Function
