Attribute VB_Name = "ListObjectHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers.ListObject"
Option Explicit

'@Description "Returns a Collection containing all the ListObjects in the specified Workbook. The Name property of each ListObject is used as the Key in the Collection. Returns an empty Collection if there are no ListObjects."
Public Function GetAllListObjects(ByVal Workbook As Workbook) As Collection
Attribute GetAllListObjects.VB_Description = "Returns a Collection containing all the ListObjects in the specified Workbook. The Name property of each ListObject is used as the Key in the Collection. Returns an empty Collection if there are no ListObjects."
    Set GetAllListObjects = New Collection
    If Workbook Is Nothing Then Exit Function
    
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        Dim ListObject As ListObject
        For Each ListObject In Worksheet.ListObjects
            GetAllListObjects.Add Item:=ListObject, Key:=ListObject.Name
        Next ListObject
    Next Worksheet
End Function

'@Description "Returns True if a ListObject with the specified Name exists in a Collection of ListObjects, and sets the provided variable to the ListObject object. Returns False if nothing is found."
Public Function TryGetListObjectFromCollection(ByVal ListObjectName As String, ByVal TableCollection As Collection, _
    ByRef OutListObject As ListObject) As Boolean
Attribute TryGetListObjectFromCollection.VB_Description = "Returns True if a ListObject with the specified Name exists in a Collection of ListObjects, and sets the provided variable to the ListObject object. Returns False if nothing is found."
    Dim ListObject As ListObject
    For Each ListObject In TableCollection
        If ListObjectName = ListObject.Name Then
            Set OutListObject = ListObject
            TryGetListObjectFromCollection = True
            Exit Function
        End If
    Next ListObject
End Function

'@Description "Returns True if any of the cells in the ListObject's Range are Locked and the Worksheet is Protected, or if the Workbook is opened in Protected Viewing mode."
Public Function IsListObjectProtected(ByVal ListObject As ListObject) As Boolean
Attribute IsListObjectProtected.VB_Description = "Returns True if any of the cells in the ListObject's Range are Locked and the Worksheet is Protected, or if the Workbook is opened in Protected Viewing mode."
    Dim Worksheet As Worksheet
    Set Worksheet = ListObject.Parent
    
    Dim Workbook As Workbook
    Set Workbook = Worksheet.Parent
    
    If ListObject.Range.Locked And Worksheet.ProtectContents Then
        IsListObjectProtected = True
    ElseIf WorkbookHelpers.IsWorkbookProtectedView(Workbook.Name) Then
        IsListObjectProtected = True
    End If
End Function

'@Description "Returns True if the Selection object contains a ListObject and sets the provided variable to the ListObject. Returns False if nothing is found."
Public Function TryGetSelectedListObject(ByRef OutListObject As ListObject) As Boolean
Attribute TryGetSelectedListObject.VB_Description = "Returns True if the Selection object contains a ListObject and sets the provided variable to the ListObject. Returns False if nothing is found."
    If Selection Is Nothing Then Exit Function
    If Not TypeOf Selection Is Range Then Exit Function
    
    Dim Range As Range
    Set Range = Selection
    If Range.ListObject Is Nothing Then Exit Function
    
    Set OutListObject = Range.ListObject
    TryGetSelectedListObject = True
End Function

'@Description "Returns True if there is exactly one ListObject in the ActiveSheet and sets the provided variable to the ListObject. Returns False if there are zero or more than one ListObjects."
Public Function TryGetSingleListObjectInActiveSheet(ByRef OutListObject As ListObject) As Boolean
Attribute TryGetSingleListObjectInActiveSheet.VB_Description = "Returns True if there is exactly one ListObject in the ActiveSheet and sets the provided variable to the ListObject. Returns False if there are zero or more than one ListObjects."
    If Application.ActiveSheet Is Nothing Then Exit Function
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Exit Function
    Dim Worksheet As Worksheet
    Set Worksheet = Application.ActiveSheet
    If Worksheet.ListObjects.Count <> 1 Then Exit Function
    
    Set OutListObject = Worksheet.ListObjects.Item(1)
    TryGetSingleListObjectInActiveSheet = True
End Function

'@Description "Returns True if the Worksheet contains a ListObject with the specified name and sets the provided variable to the ListObject object. Returns False if nothing is found."
Public Function TryGetListObjectInWorksheet(ByVal ListObjectName As String, ByVal Worksheet As Worksheet, _
    ByRef OutListObject As ListObject) As Boolean
Attribute TryGetListObjectInWorksheet.VB_Description = "Returns True if the Worksheet contains a ListObject with the specified name and sets the provided variable to the ListObject object. Returns False if nothing is found."
    If Worksheet Is Nothing Then Exit Function
    If ListObjectName = vbNullString Then Exit Function
    
    Dim ListObject As ListObject
    For Each ListObject In Worksheet.ListObjects
        If ListObject.Name = ListObjectName Then
            Set OutListObject = ListObject
            TryGetListObjectInWorksheet = True
            Exit Function
        End If
    Next ListObject
End Function
