Attribute VB_Name = "PathHelpers"
'@IgnoreModule AssignmentNotUsed
'@Folder "Version4.ViewModels"
Option Explicit

Private Const CHR_SEPARATOR As String = "\"

Public Function CreatePathTreeItems(ByVal InputArray As Variant) As Variant
    Debug.Assert IsArray(InputArray)
    Debug.Assert LBound(InputArray) = 1
    Debug.Assert UBound(InputArray) > 1
    
    ArraySort.QuickSort InputArray

    Dim InputArrayUnique As Variant
    InputArrayUnique = ArrayUnique.Unique(InputArray)
    Debug.Assert UBound(InputArrayUnique) > 1
    
    Dim Results As Variant
    ReDim Results(1 To UBound(InputArrayUnique))
    
    Dim Cursor As Long
    
    Dim i As Long
    For i = 1 To UBound(InputArrayUnique)
        Dim Result As String
        Result = InputArrayUnique(i)
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
    PreviousSeparator = InStrRev(Path, CHR_SEPARATOR)
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
    IsNodePath = (Left$(Key, Len(NODE_PREFIX_PATH)) = NODE_PREFIX_PATH)
End Function
