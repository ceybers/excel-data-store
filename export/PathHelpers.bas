Attribute VB_Name = "PathHelpers"
'@Folder "Version062.ViewModels"
Option Explicit

Public Function CreatePathTreeItems(ByVal InputArray As Variant) As Variant
    Debug.Assert IsArray(InputArray)
    Debug.Assert LBound(InputArray) = 1
    Debug.Assert UBound(InputArray) > 1
    
    ArraySort.QuickSort InputArray
    InputArray = ArrayUnique.Unique(InputArray)
    Debug.Assert UBound(InputArray) > 1
    
    Dim Results As Variant
    ReDim Results(1 To UBound(InputArray))
    
    Dim Cursor As Long
    
    Dim i As Long
    For i = 1 To UBound(InputArray)
        Dim Result As String
        Result = InputArray(i)
        Do While Result <> vbNullString
            Cursor = Cursor + 1
            Results(Cursor) = Result
            If Cursor = UBound(Results) Then
                ReDim Preserve Results(1 To Cursor * 2)
            End If
            Result = RecursePathParent(Result)
        Loop
    Next i
    
    ReDim Preserve Results(1 To Cursor)
    ArraySort.QuickSort Results
    Results = ArrayUnique.Unique(Results)
    
    CreatePathTreeItems = Results
End Function

Private Function RecursePathParent(ByVal Path As String) As String
    Dim PreviousSeparator As Long
    PreviousSeparator = InStrRev(Path, "\")
    If PreviousSeparator = 0 Then
        RecursePathParent = vbNullString
    Else
        RecursePathParent = Left$(Path, PreviousSeparator - 1)
    End If
End Function

Public Function PrefixNode(ByVal Key As String) As String
    PrefixNode = NODE_PREFIX_PATH & Key
End Function

Public Function IsNodePath(ByVal Key As String) As Boolean
    IsNodePath = (Left$(Key, 5) = NODE_PREFIX_PATH)
End Function
