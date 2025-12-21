Attribute VB_Name = "WorksheetHelpers"
'@Folder "Helpers.Worksheet"
Option Explicit

'@Description "Returns the range of cells in a Worksheet starting from the first row beneath the header row ranging until the last row in the UsedRange. Assumes header row is always Row 1. Returns Nothing if there are no rows, or if there is only a header row."
Public Function GetWorksheetDatabodyRange(ByVal Worksheet As Worksheet) As Range
Attribute GetWorksheetDatabodyRange.VB_Description = "Returns the range of cells in a Worksheet starting from the first row beneath the header row ranging until the last row in the UsedRange. Assumes header row is always Row 1. Returns Nothing if there are no rows, or if there is only a header row."
    If Worksheet Is Nothing Then Exit Function
    If Worksheet.UsedRange Is Nothing Then Exit Function
    If Worksheet.UsedRange.Rows.Count = 1 Then Exit Function
    
    Set GetWorksheetDatabodyRange = Worksheet.UsedRange.Offset(1, 0).Resize(Worksheet.UsedRange.Rows.Count - 1)
End Function

'@Description "Returns True if the Worksheet with the specified name exists in the specified Workbook and sets the provided variable to the Worksheet object. Returns False if nothing is found."
Public Function TryGetWorksheet(ByVal WorksheetName As String, ByVal Workbook As Workbook, _
    ByRef OutWorksheet As Worksheet) As Boolean
Attribute TryGetWorksheet.VB_Description = "Returns True if the Worksheet with the specified name exists in the specified Workbook and sets the provided variable to the Worksheet object. Returns False if nothing is found."
    If Workbook Is Nothing Then Exit Function
    If WorksheetName = vbNullString Then Exit Function
    
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        If Worksheet.Name = WorksheetName Then
            Set OutWorksheet = Worksheet
            TryGetWorksheet = True
            Exit Function
        End If
    Next Worksheet
End Function
