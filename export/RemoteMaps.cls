VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteMaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "RemoteDataStore.Models.Plural"
Option Explicit

'Private Const COL_ID_NAME As String = "MapID"
Private Const ARRAY_WIDTH As Long = 6
Private Const HASH_LEN  As Long = 4

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

Public Property Get Item(ByVal Index As Long) As RemoteMap
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
    
    Dim WorksheetValues As Variant
    WorksheetValues = DataBodyRange.Value2
    
    Dim i As Long
    For i = 1 To UBound(WorksheetValues, 1)
        Dim NewRemoteMap As RemoteMap
        Set NewRemoteMap = New RemoteMap
        NewRemoteMap.LoadFromArray WorksheetValues, i
        This.Item.Add Item:=NewRemoteMap
    Next i
End Sub

Public Sub Save()
    Debug.Assert Not This.Worksheet Is Nothing
    
    If This.IsDirty = False Then Exit Sub
    
    ClearWorksheet
    
    Dim UpdatedValues As Variant
    ReDim UpdatedValues(1 To Count, 1 To ARRAY_WIDTH)
    
    Dim i As Long
    For i = 1 To Count
        Dim ThisRemoteMap As RemoteMap
        Set ThisRemoteMap = This.Item.Item(i)
        ThisRemoteMap.SaveToArray UpdatedValues, i
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
        .Item(1).NumberFormat = NUMFMT_TEXT ' MapID
        .Item(4).NumberFormat = NUMFMT_DATE ' CTime
        .Item(5).NumberFormat = NUMFMT_DATE ' MTime
        .Item(6).NumberFormat = NUMFMT_DATE ' DTime
    End With
End Sub

Private Sub SortWorksheet()
    This.Worksheet.UsedRange.Sort Key1:=This.Worksheet.Columns.Item(1), Header:=xlYes
End Sub

Public Function AddOrGet(ByVal SerializedString As String, ByVal Caption As String) As RemoteMap
    Dim NewID As String
    NewID = Left$(HashSHA1(SerializedString & Chr$(ASCII_US) & Caption), HASH_LEN)
    
    Dim RemoteMap As RemoteMap
    If TryGetByID(NewID, RemoteMap) Then
        Set AddOrGet = RemoteMap
        Exit Function
    Else
        Set AddOrGet = Add(SerializedString, Caption)
    End If
End Function

Public Function Add(ByVal SerializedString As String, ByVal Caption As String) As RemoteMap
    Dim NewRemoteMap As RemoteMap
    Set NewRemoteMap = New RemoteMap
    
    Dim NewID As String
    NewID = Left$(HashSHA1(SerializedString & Chr$(ASCII_US) & Caption), HASH_LEN)
    
    If TryGetByID(NewID, Nothing) Then Exit Function
    
    With NewRemoteMap
        .ID = NewID
        .SerializedString = SerializedString
        .Caption = Caption
        .CreationTime = Application.Evaluate("=NOW()")
    End With
    
    This.Item.Add NewRemoteMap
    This.IsDirty = True
    
    Set Add = NewRemoteMap
End Function

Public Function TryGetByID(ByVal ID As String, ByRef OutRemoteMap As RemoteMap) As Boolean
    Dim ThisRemoteMap As RemoteMap
    For Each ThisRemoteMap In This.Item
        If ThisRemoteMap.ID = ID Then
            Set OutRemoteMap = ThisRemoteMap
            TryGetByID = True
            Exit Function
        End If
    Next ThisRemoteMap
End Function
