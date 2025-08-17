VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableMapMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TableMapMatcher.Model"
Option Explicit

Private Type TTableMapMatch
    Caption As String
    Timestamp As Date
    SerializedTable As SerializedTable
    TableMap As TableMap
    ListObject As ListObject
    Score As Long
    Results As Variant
    FieldCount As Long
    MatchCount As Long
End Type
Private This As TTableMapMatch

Public Property Get Caption() As String
    Caption = This.Caption
End Property

Public Property Get Timestamp() As Date
    Timestamp = This.Timestamp
End Property

Public Property Get Score() As Long
    Score = This.Score
End Property

Public Property Get Results() As Variant
    Results = This.Results
End Property

Public Sub Load(ByVal RemoteTable As RemoteTable, ByVal RemoteMap As RemoteMap)
    This.Caption = RemoteTable.Caption
    This.Timestamp = RemoteTable.ModificationTime
    
    Set This.SerializedTable = New SerializedTable
    This.SerializedTable.Deserialize RemoteTable.SerializedString
    
    Set This.TableMap = New TableMap
    This.TableMap.Deserialize RemoteMap.SerializedString
End Sub

Public Sub Evaluate(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    This.FieldCount = This.TableMap.FieldMap.Count
    UpdateResults
    UpdateScore
End Sub

Private Sub UpdateResults()
    Dim Worksheet As Worksheet
    Set Worksheet = This.ListObject.Parent
    
    Dim Workbook As Workbook
    Set Workbook = Worksheet.Parent
    
    Dim Results As Variant
    ReDim Results(1 To 7 + This.FieldCount, 1 To 3)
    
    Results(1, 1) = "FullName"
    Results(1, 2) = This.SerializedTable.WorkbookFullName
    Results(1, 3) = This.SerializedTable.WorkbookFullName = Workbook.FullName
    
    Results(2, 1) = "Path"
    Results(2, 2) = This.SerializedTable.WorkbookPath
    Results(2, 3) = This.SerializedTable.WorkbookPath = Workbook.Path
    
    Results(3, 1) = "Filename"
    Results(3, 2) = This.SerializedTable.WorkbookName
    Results(3, 3) = This.SerializedTable.WorkbookName = Workbook.Name
    
    Results(4, 1) = "Worksheet"
    Results(4, 2) = This.SerializedTable.WorksheetName
    Results(4, 3) = This.SerializedTable.WorksheetName = Worksheet.Name
    
    Results(5, 1) = "Table"
    Results(5, 2) = This.SerializedTable.ListObjectName
    Results(5, 3) = This.SerializedTable.ListObjectName = This.ListObject.Name
    
    Results(6, 1) = "Key"
    Results(6, 2) = This.TableMap.KeyMap.KeyColumnName
    Results(6, 3) = ListColumnExists(This.TableMap.KeyMap.KeyColumnName, This.ListObject)
    
    Results(7, 1) = "Fields: "
    
    This.MatchCount = 0
    
    Dim i As Long
    For i = 1 To This.FieldCount
        Results(7 + i, 1) = "Field #" & i
        Results(7 + i, 2) = This.TableMap.FieldMap.Item(i).ListColumnName
        
        Dim FieldMatch As Boolean
        FieldMatch = ListColumnExists(Results(7 + i, 2), This.ListObject)
        
        Results(7 + i, 3) = FieldMatch
        If FieldMatch Then This.MatchCount = This.MatchCount + 1
    Next i
    
    If This.FieldCount = 0 Then
        Results(7, 3) = "No Fields"
    ElseIf This.MatchCount = This.FieldCount Then
        Results(7, 3) = "All (" & This.MatchCount & " of " & This.FieldCount & ")"
    ElseIf This.MatchCount > 0 Then
        Results(7, 3) = "Some (" & This.MatchCount & " of " & This.FieldCount & ")"
    Else
        Results(7, 3) = "None (" & This.MatchCount & " of " & This.FieldCount & ")"
    End If
    
    This.Results = Results
End Sub

Private Sub UpdateScore()
    Dim Score As Long
    Score = 0
    Score = IIf(This.Results(1, 3), Score + 1000, 0)
    Score = IIf(This.Results(3, 3), Score + 500, 0)
    Score = IIf(This.Results(2, 3), Score + 250, 0)
    Score = IIf(This.Results(4, 3), Score + 500, 0)
    Score = IIf(This.Results(5, 3), Score + 500, 0)
    
    If This.FieldCount > 0 Then
        Score = Score + ((This.MatchCount / This.FieldCount) * 1000)
    End If
    Score = IIf(This.Results(6, 3), Score, 0)
    
    This.Score = Score
End Sub
