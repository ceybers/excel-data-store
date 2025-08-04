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
