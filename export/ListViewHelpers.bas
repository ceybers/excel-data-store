Attribute VB_Name = "ListViewHelpers"
'@Folder "Helpers.Controls"
Option Explicit

Public Sub ListViewColumnClickSort(ByVal ListView As ListView, ByVal Index As Long)
    If ListView.SortKey = Index - 1 Then
        ListView.SortOrder = IIf(ListView.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    End If
    
    ListView.SortKey = Index - 1
End Sub
