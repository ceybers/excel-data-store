Attribute VB_Name = "WorkbookHelpers"
'@Folder "Helpers.Workbook"
Option Explicit

'@Description "Returns True if a Workbook with the specified name is open in Excel and sets the provided variable to the Workbook object. Returns False if nothing is found."
Public Function TryGetWorkbook(ByVal WorkbookName As String, ByRef OutWorkbook As Workbook) As Boolean
Attribute TryGetWorkbook.VB_Description = "Returns True if a Workbook with the specified name is open in Excel and sets the provided variable to the Workbook object. Returns False if nothing is found."
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

'@Description "Returns True if the specified Workbook is opened in Protected View. Returns False if it is not, or if the variable is set to Nothing."
Public Function IsWorkbookProtectedView(ByVal WorkbookName As String) As Boolean
Attribute IsWorkbookProtectedView.VB_Description = "Returns True if the specified Workbook is opened in Protected View. Returns False if it is not, or if the variable is set to Nothing."
    If WorkbookName = vbNullString Then Exit Function
    
    Dim ProtectedViewWindow As ProtectedViewWindow
    For Each ProtectedViewWindow In Application.ProtectedViewWindows
        If ProtectedViewWindow.Workbook.Name = WorkbookName Then
            IsWorkbookProtectedView = True
            Exit Function
        End If
    Next ProtectedViewWindow
End Function

'@Description "Returns True if the specified variable is referencing a Workbook that is still open. Returns False if the variable is referencing a Workbook that has been closed, or if the variable is set to Nothing."
Public Function IsWorkbookOpen(ByVal Workbook As Workbook) As Boolean
Attribute IsWorkbookOpen.VB_Description = "Returns True if the specified variable is referencing a Workbook that is still open. Returns False if the variable is referencing a Workbook that has been closed, or if the variable is set to Nothing."
    If Workbook Is Nothing Then Exit Function
    
    Dim TestWorkbook As String
    On Error GoTo ErrorInIsWorkbookOpen
    TestWorkbook = Workbook.Name
    On Error Resume Next
    
    IsWorkbookOpen = True
    Exit Function
    
ErrorInIsWorkbookOpen:
    Exit Function
End Function
