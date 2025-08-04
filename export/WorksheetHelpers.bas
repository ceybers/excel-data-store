Attribute VB_Name = "WorksheetHelpers"
'@Folder "Helpers.Worksheet"
Option Explicit

'@Description "Returns the range of cells in a Worksheet starting from the first row beneath the header row ranging until the last row in the Used Range. Assumes header row is always Row 1. Returns Nothing if there are no rows."
Public Function GetWorksheetDatabodyRange(ByVal Worksheet As Worksheet) As Range
Attribute GetWorksheetDatabodyRange.VB_Description = "Returns the range of cells in a Worksheet starting from the first row beneath the header row ranging until the last row in the Used Range. Assumes header row is always Row 1. Returns Nothing if there are no rows."
    If Worksheet Is Nothing Then Exit Function
    If Worksheet.UsedRange Is Nothing Then Exit Function
    If Worksheet.UsedRange.Rows.Count = 1 Then Exit Function
    
    Set GetWorksheetDatabodyRange = Worksheet.UsedRange.Offset(1, 0).Resize(Worksheet.UsedRange.Rows.Count - 1)
End Function
