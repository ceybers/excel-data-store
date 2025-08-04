Attribute VB_Name = "modTEST4"
'@IgnoreModule ProcedureNotUsed
'@Folder("Version4")
Option Explicit

'@EntryPoint
Public Sub TESTClearTable()
    Dim ListObject As ListObject
    Set ListObject = TESTListObject
    Dim RangeToClear As Range
    Set RangeToClear = ListObject.DataBodyRange.Offset(0, 1)
    RangeToClear.ClearContents
End Sub

Private Function TESTListObject() As ListObject
    If Not Selection.ListObject Is Nothing Then
        Set TESTListObject = Selection.ListObject
        Exit Function
    End If
    
    If Selection.Parent.ListObjects.Count = 1 Then
        Set TESTListObject = Selection.Parent.ListObjects.Item(1)
        Exit Function
    End If
    
    Set TESTListObject = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
End Function

Private Function ZZZTestTableMap() As TableMap
    Dim Result As TableMap
    Set Result = New TableMap
    'Result.Deserialize "MyKey¤Test\NATO§DeletionTime¤575E¥CreationTime¤324F¥ModifiedTime¤A2AC"
    'Result.Deserialize "Airport Name¤Test\Airports§Airport Location¤6252¥Airport Code¤F232¥Total Passengers¤DFF6"
    Set ZZZTestTableMap = Result
End Function

Private Sub ZZZTableMapSerialization()
    Dim KeyMap As KeyMap
    Set KeyMap = New KeyMap
    With KeyMap
        .KeyColumnName = "MyKey"
        .KeyPath = "Test\NATO"
        Debug.Print .Serialize; vbCrLf
    End With
    
    Dim FieldMap As FieldMap
    Set FieldMap = New FieldMap
    With FieldMap
        .Add "Foo", "ABCD"
        .Add "Bar", "BCDE"
        .Add "Baz", "CDEF"
        Debug.Print .Serialize; vbCrLf
    End With
    
    Dim TableMap As TableMap
    Set TableMap = New TableMap
    With TableMap
        .Deserialize "MyKey¤Test\NATO§DeletionTime¤575E¥CreationTime¤324F¥ModifiedTime¤A2AC"
        .KeyMap.DEBUGPrint
        .FieldMap.DEBUGPrint
    End With
    
    Dim LO As ListObject
    Set LO = Selection.ListObject
    With LO
        Debug.Print "FullName = "; LO.Parent.Parent.FullName
        Debug.Print "Path     = "; LO.Parent.Parent.Path
        Debug.Print "WBName   = "; LO.Parent.Parent.Name
        Debug.Print "WSName   = "; LO.Parent.Name
        Debug.Print "LOName   = "; LO.Name
    End With
End Sub
