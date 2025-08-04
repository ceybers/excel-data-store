VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteCommits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder("RemoteDataStore.Models.Plural")
Option Explicit

Private Const DO_DEBUG As Boolean = False
Private Const ARRAY_WIDTH As Long = 6
Private Const HASH_LEN  As Long = 40

Private Type TState
    Worksheet As Worksheet
    Item As Collection
    IsDirty As Boolean
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Item = New Collection
End Sub

Public Property Get Count() As Long
    Count = This.Item.Count
End Property

Public Property Get Item(ByVal Index As Long) As RemoteCommit
    Set Item = This.Item.Item(Index)
End Property

Public Sub Load(ByVal Worksheet As Worksheet)
    Set This.Worksheet = Worksheet
    
    LoadFromWorksheet
End Sub

Private Sub LoadFromWorksheet()
    SortWorksheet
    
    Dim DataBodyRange As Range
    Set DataBodyRange = GetWorksheetDatabodyRange(This.Worksheet)
    If DataBodyRange Is Nothing Then Exit Sub
    
    Dim ExistingValues As Variant
    If DO_DEBUG Then Debug.Print "RemoteCommits.LoadFromWorksheet reading from " & DataBodyRange.Address(False, False)
    ExistingValues = DataBodyRange.Value2
        
    Dim i As Long
    For i = 1 To UBound(ExistingValues, 1)
        Dim NewRemoteCommit As RemoteCommit
        Set NewRemoteCommit = New RemoteCommit
        NewRemoteCommit.LoadFromArray ExistingValues, i
        This.Item.Add Item:=NewRemoteCommit
    Next i
End Sub

Public Sub Save()
    Debug.Assert Not This.Worksheet Is Nothing
    
    If This.IsDirty = False Then Exit Sub
    
    ClearWorksheet
    
    Dim UpdatedValues As Variant
    ReDim UpdatedValues(1 To This.Item.Count, 1 To ARRAY_WIDTH)
    
    Dim i As Long
    For i = 1 To This.Item.Count
        Dim ThisRemoteCommit As RemoteCommit
        Set ThisRemoteCommit = This.Item.Item(i)
        ThisRemoteCommit.SaveToArray UpdatedValues, i
    Next i
    
    SetNumberFormats
    
    RangeSetValueFromVariant This.Worksheet.Range("A2"), UpdatedValues
    
    SortWorksheet
End Sub

Private Sub ClearWorksheet()
    This.Worksheet.AutoFilterMode = False
    If This.Worksheet.UsedRange Is Nothing Then Exit Sub
    If This.Worksheet.UsedRange.Rows.Count = 1 Then Exit Sub
    This.Worksheet.UsedRange.Offset(1).Delete
End Sub

Private Sub SetNumberFormats()
    With This.Worksheet.Columns
        .Item(1).NumberFormat = NUMFMT_TEXT ' CommitID
        .Item(2).NumberFormat = NUMFMT_TEXT ' TableID
        .Item(3).NumberFormat = NUMFMT_TEXT ' MapID
        .Item(4).NumberFormat = NUMFMT_TEXT ' Caption
        .Item(5).NumberFormat = NUMFMT_DATE ' Timestamp
        .Item(6).NumberFormat = NUMFMT_TEXT ' Message
    End With
End Sub

Private Sub SortWorksheet()
    This.Worksheet.UsedRange.Sort Key1:=This.Worksheet.Columns.Item(1), Header:=xlYes
End Sub

Public Function Add(ByVal TableID As String, ByVal MapID As String, ByVal Caption As String) As RemoteCommit
    Dim NewRemoteCommit As RemoteCommit
    Set NewRemoteCommit = CreateNewCommit(TableID, MapID, Caption)
        
    This.Item.Add NewRemoteCommit
    This.IsDirty = True
    
    Save
    
    Set Add = NewRemoteCommit
End Function

Private Function CreateNewCommit(ByVal TableID As String, ByVal MapID As String, ByVal Caption As String) As RemoteCommit
    Dim NewRemoteCommit As RemoteCommit
    Set NewRemoteCommit = New RemoteCommit
    
    With NewRemoteCommit
        .TableID = TableID
        .MapID = MapID
        .Caption = Caption & " (" & Format$(Now(), "yyyy/mm/dd hh:MM:ss") & ")"
        .Timestamp = Application.Evaluate("=NOW()")
        .CommitID = Left$(HashSHA1(TableID & Chr$(ASCII_US) & MapID & Chr$(ASCII_US) & .Timestamp), HASH_LEN)
    End With
    
    Set CreateNewCommit = NewRemoteCommit
End Function
