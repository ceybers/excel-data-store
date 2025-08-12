Attribute VB_Name = "WorkbookHelpers"
'@Folder "Helpers.Workbook"
Option Explicit

'@Description "Tries to return the Workbook with the given name if it is currently open in this instance of Excel."
Public Function TryGetWorkbook(ByVal WorkbookName As String, ByRef OutWorkbook As Workbook) As Boolean
Attribute TryGetWorkbook.VB_Description = "Tries to return the Workbook with the given name if it is currently open in this instance of Excel."
    If WorkbookName = vbNullString Then Exit Function
    
    Dim Workbook As Workbook
    For Each Workbook In Application.Workbooks
        If Workbook.Name = WorkbookName Then
            Set OutWorkbook = Workbook
            TryGetWorkbook = True
            Exit Function
        End If
    Next Workbook
End Function

'@Description "Returns True if the given Workbook is opened in Protected View"
Public Function IsWorkbookProtectedView(ByVal WorkbookName As String) As Boolean
Attribute IsWorkbookProtectedView.VB_Description = "Returns True if the given Workbook is opened in Protected View"
    If WorkbookName = vbNullString Then Exit Function
    
    Dim ProtectedViewWindow As ProtectedViewWindow
    For Each ProtectedViewWindow In Application.ProtectedViewWindows
        If ProtectedViewWindow.Workbook.Name = WorkbookName Then
            IsWorkbookProtectedView = True
            Exit Function
        End If
    Next ProtectedViewWindow
End Function

'@Description "Returns True if a Workbook is still open. Returns False if the workbook is closed but the reference is still present."
Public Function IsWorkbookOpen(ByVal Workbook As Workbook) As Boolean
    If Workbook Is Nothing Then Exit Function
    Dim TestWorkbook As String
    On Error Resume Next
    TestWorkbook = Workbook.Name
    On Error GoTo 0
    If TestWorkbook = vbNullString Then Exit Function
    
    IsWorkbookOpen = True
End Function
