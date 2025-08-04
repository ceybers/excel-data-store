VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "RemoteDataStore.Models.Plural"
Option Explicit

Private Const ARRAY_WIDTH As Long = 9
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

Public Property Get Item(ByVal Index As Long) As RemoteTable
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
        Dim NewRemoteTable As RemoteTable
        Set NewRemoteTable = New RemoteTable
        NewRemoteTable.LoadFromArray WorksheetValues, i
        This.Item.Add Item:=NewRemoteTable
    Next i
End Sub

Public Sub Save()
    Debug.Assert Not This.Worksheet Is Nothing
    
    If This.IsDirty = False Then Exit Sub
    
    ClearWorksheet
    
    Dim UpdateValues As Variant
    ReDim UpdateValues(1 To Count, 1 To ARRAY_WIDTH)
    
    Dim i As Long
    For i = 1 To Count
        Dim ThisRemoteTable As RemoteTable
        Set ThisRemoteTable = This.Item.Item(i)
        ThisRemoteTable.SaveToArray UpdateValues, i
    Next i
    
    SetNumberFormats
    
    RangeSetValueFromVariant This.Worksheet.Range("A2"), UpdateValues
    
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
        .Item(1).NumberFormat = NUMFMT_TEXT ' TableID
        .Item(4).NumberFormat = NUMFMT_DATE ' CTime
        .Item(5).NumberFormat = NUMFMT_DATE ' MTime
        .Item(6).NumberFormat = NUMFMT_DATE ' DTime
        .Item(7).NumberFormat = NUMFMT_DATE ' LastPull
        .Item(8).NumberFormat = NUMFMT_DATE ' LastPush
    End With
End Sub

Private Sub SortWorksheet()
    This.Worksheet.UsedRange.Sort Key1:=This.Worksheet.Columns.Item(1), Header:=xlYes
End Sub

Public Function AddOrGet(ByVal SerializedString As String, ByVal Caption As String, ByVal MapID As String) As RemoteTable
    Dim NewID As String
    NewID = Left$(HashSHA1(SerializedString & Chr$(ASCII_US) & Caption & Chr$(ASCII_US) & MapID), HASH_LEN)
    
    Dim RemoteTable As RemoteTable
    If TryGetByID(NewID, RemoteTable) Then
        Set AddOrGet = RemoteTable
        Exit Function
    Else
        Set AddOrGet = Add(SerializedString, Caption, MapID)
    End If
End Function

Public Function Add(ByVal SerializedString As String, ByVal Caption As String, ByVal MapID As String) As RemoteTable
    Dim NewRemoteTable As RemoteTable
    Set NewRemoteTable = New RemoteTable
    
    Dim NewID As String
    NewID = Left$(HashSHA1(SerializedString & Chr$(ASCII_US) & Caption & Chr$(ASCII_US) & MapID), HASH_LEN)
    
    If TryGetByID(NewID, Nothing) Then Exit Function
    
    With NewRemoteTable
        .ID = NewID
        .SerializedString = SerializedString
        .Caption = Caption
        .MapID = MapID
        .CreationTime = Application.Evaluate("=NOW()")
    End With
    
    This.Item.Add NewRemoteTable
    This.IsDirty = True
    
    Set Add = NewRemoteTable
End Function

Private Function TryGetByID(ByVal ID As String, ByRef OutRemoteTable As RemoteTable) As Boolean
    Dim ThisRemoteTable As RemoteTable
    For Each ThisRemoteTable In This.Item
        If ThisRemoteTable.ID = ID Then
            Set OutRemoteTable = ThisRemoteTable
            TryGetByID = True
            Exit Function
        End If
    Next ThisRemoteTable
End Function
