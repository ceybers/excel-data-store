Attribute VB_Name = "Module1"
'@Folder("Helpers.ListObject")
Option Explicit

Public Sub aaa()
    Dim lo As ListObject
    If TryGetSelectedListObject(lo) = True Then
        Debug.Print "T="; lo.Name
    Else
        Debug.Print "F"
    End If
    
        If TryGetActiveSheetListObject(lo) = True Then
        Debug.Print "T="; lo.Name
    Else
        Debug.Print "F"
    End If
    
    Dim r As Range
    If TryGetSelectionRange(r) Then
        Debug.Print r.Address
    Else
        Debug.Print "no r"
    End If
End Sub
