Attribute VB_Name = "ChangeMask"
'@Folder("Version4.Queries")
Option Explicit

Public Function GetChangeMask(ByVal LocalValues As Variant, ByVal RemoteValues As Variant) As Variant
    Dim Result As Variant
    ReDim Result(1 To UBound(LocalValues, 1), 1 To UBound(LocalValues, 2)) As Boolean
    
    Dim i As Long
    For i = 1 To UBound(LocalValues, 1)
        Dim j As Long
        For j = 1 To UBound(LocalValues, 2)
            Result(i, j) = LocalValues(i, j) <> RemoteValues(i, j)
        Next j
    Next i
    
    GetChangeMask = Result
End Function

Public Function CountChanges(ByVal ChangeMask As Variant) As Long
    Dim Result As Long
    
    Dim i As Long
    For i = 1 To UBound(ChangeMask, 1)
        Dim j As Long
        For j = 1 To UBound(ChangeMask, 2)
            If ChangeMask(i, j) = True Then
                Result = Result + 1
            End If
        Next j
    Next i
    
    CountChanges = Result
End Function
