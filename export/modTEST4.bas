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

'@EntryPoint
Public Sub TESTRandomiseTable()
    Dim ListObject As ListObject
    Set ListObject = TESTListObject
    Dim RangeToRandomise As Range
    Set RangeToRandomise = ListObject.DataBodyRange.Offset(0, 1).Resize(ListObject.DataBodyRange.Rows.Count, ListObject.ListColumns.Count - 1)
    
    RangeToRandomise.Formula2 = "=RAND()"
    RangeToRandomise.Value2 = RangeToRandomise.Value2
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

Public Sub AAA()
    ThisWorkbook.Worksheets.Item("Sheet1").Range("B3:C3,C9:C10,C5:D7,F11").Select
    
    Dim DataBodyRange As Range
    Set DataBodyRange = ThisWorkbook.Worksheets.Item("Sheet1").Range("C2:C11")
    
    Dim vv2 As Variant
    vv2 = GetStaggeredArrayValues(DataBodyRange, Selection)
    
    Stop
End Sub

