Attribute VB_Name = "ListObjectHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers.ListObject"
Option Explicit

'@Description "Returns a Collection containing all the ListObjects in a given Workbook"
Public Function GetAllListObjects(ByVal Workbook As Workbook) As Collection
Attribute GetAllListObjects.VB_Description = "Returns a Collection containing all the ListObjects in a given Workbook"
    Set GetAllListObjects = New Collection
    
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        Dim ListObject As ListObject
        For Each ListObject In Worksheet.ListObjects
            GetAllListObjects.Add Item:=ListObject, Key:=ListObject.Name
        Next ListObject
    Next Worksheet
End Function

'@Description "Tries to return the ListObject with the given name from a Collection of ListObjects."
Public Function TryGetListObjectFromCollection(ByVal TableCollection As Collection, ByVal ListObjectName As String, ByRef OutListObject As ListObject) As Boolean
Attribute TryGetListObjectFromCollection.VB_Description = "Tries to return the ListObject with the given name from a Collection of ListObjects."
    Dim ListObject As ListObject
    For Each ListObject In TableCollection
        If ListObjectName = ListObject.Name Then
            Set OutListObject = ListObject
            TryGetListObjectFromCollection = True
            Exit Function
        End If
    Next ListObject
End Function

'@Description "Returns True if any of the cells in ListObject Range are Locked and the Worksheet is Protected, or if the Workbook is opened in Protected Viewing mode."
Public Function TestIfProtected(ByVal ListObject As ListObject) As Boolean
Attribute TestIfProtected.VB_Description = "Returns True if any of the cells in ListObject Range are Locked and the Worksheet is Protected, or if the Workbook is opened in Protected Viewing mode."
    Dim Worksheet As Worksheet
    Set Worksheet = ListObject.Parent
    
    Dim Workbook As Workbook
    Set Workbook = Worksheet.Parent
    
    If ListObject.Range.Locked And Worksheet.ProtectContents Then
        TestIfProtected = True
    ElseIf WorkbookHelpers.IsWorkbookProtectedView(Workbook.Name) Then
        TestIfProtected = True
    End If
End Function
