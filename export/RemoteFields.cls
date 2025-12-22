VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "RemoteDataStore.Models.Plural"
Option Explicit

Private Const ARRAY_WIDTH As Long = 7

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

Public Property Get Item(ByVal Index As Long) As RemoteField
    Set Item = This.Item.Item(Index)
End Property

Public Sub Load(ByVal Worksheet As Worksheet)
    Set This.Worksheet = Worksheet
    
    GenerateMetaFieldObjects
    LoadFromWorksheet
End Sub

Private Sub GenerateMetaFieldObjects()
    With This.Item
        .Add Item:=MetaRemoteFields.Unmapped
        .Add Item:=MetaRemoteFields.AddNew
    End With
End Sub

Private Sub LoadFromWorksheet()
    SortWorksheet
    
    Dim DataBodyRange As Range
    Set DataBodyRange = GetWorksheetDatabodyRange(This.Worksheet)
    If DataBodyRange Is Nothing Then Exit Sub
    
    Dim WorksheetValues As Variant
    Log.Message "RemoteFields.LoadFromWorksheet reading from " & DataBodyRange.Address(False, False), "RFields", Info_Level
    WorksheetValues = DataBodyRange.Value2
    
    Dim i As Long
    For i = 1 To UBound(WorksheetValues, 1)
        Dim NewRemoteField As RemoteField
        Set NewRemoteField = New RemoteField
        NewRemoteField.LoadFromArray WorksheetValues, i
        This.Item.Add Item:=NewRemoteField
    Next i
End Sub

Public Sub Save()
    Debug.Assert Not This.Worksheet Is Nothing
    
    If This.IsDirty = False Then Exit Sub
    
    ClearWorksheet
    
    Dim ItemCount As Long
    ItemCount = Count - META_FIELD_COUNT
    
    Dim UpdatedValues As Variant
    ReDim UpdatedValues(1 To ItemCount, 1 To ARRAY_WIDTH)
    
    Dim i As Long
    For i = 1 To ItemCount
        Dim ThisRemoteField As RemoteField
        Set ThisRemoteField = This.Item.Item(i + META_FIELD_COUNT)
        ThisRemoteField.SaveToArray UpdatedValues, i
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
        .Item(1).NumberFormat = NUMFMT_TEXT ' ID
        .Item(5).NumberFormat = NUMFMT_DATE ' CTime
        .Item(6).NumberFormat = NUMFMT_DATE ' MTime
        .Item(7).NumberFormat = NUMFMT_DATE ' DTime
    End With
End Sub

Private Sub SortWorksheet()
    This.Worksheet.UsedRange.Sort Key1:=This.Worksheet.Columns.Item(1), Header:=xlYes
End Sub

Public Function GetIndexFromID(ByVal ID As String) As Long
    Dim i As Long
    For i = 1 To This.Item.Count
        Dim ThisRemoteField As RemoteField
        Set ThisRemoteField = This.Item.Item(i)
        If ThisRemoteField.ID = ID Then
            GetIndexFromID = i
            Exit Function
        End If
    Next i
    
    GetIndexFromID = -1
End Function

Public Function TryGetRemoteFieldFromID(ByVal ID As String, ByRef OutRemoteField As RemoteField) As Boolean
    Dim ThisRemoteField As RemoteField
    For Each ThisRemoteField In This.Item
        If ThisRemoteField.ID = ID Then
            Set OutRemoteField = ThisRemoteField
            TryGetRemoteFieldFromID = True
            Exit Function
        End If
    Next ThisRemoteField
End Function

Public Sub AddNew(ByVal Path As String, ByVal Name As String, ByVal Caption As String)
    If Path = vbNullString Then Exit Sub
    If Name = vbNullString Then Exit Sub
    If Caption = vbNullString Then Exit Sub
    
    Dim ExistingRemoteField As RemoteField
    If TryGetRemoteFieldFromPathAndName(Path, Name, ExistingRemoteField) Then Exit Sub
    
    Dim NewRemoteField As RemoteField
    Set NewRemoteField = CreateNewField(Path, Name, Caption)

    This.Item.Add NewRemoteField
    This.IsDirty = True
    
    Save 'Saving in the same procedure that Adds because it is called from outside the class module
End Sub

Private Function TryGetRemoteFieldFromPathAndName(ByVal Path As String, ByVal Name As String, ByRef OutRemoteField As RemoteField) As Boolean
    Dim ThisRemoteField As RemoteField
    For Each ThisRemoteField In This.Item
        If ThisRemoteField.ID <> FIELD_ID_ADDNEW And ThisRemoteField.Path = Path And ThisRemoteField.Name = Name Then
            Set OutRemoteField = ThisRemoteField
            TryGetRemoteFieldFromPathAndName = True
            Exit Function
        End If
    Next ThisRemoteField
End Function

Private Function CreateNewField(ByVal Path As String, ByVal Name As String, ByVal Caption As String) As RemoteField
    Dim NewRemoteField As RemoteField
    Set NewRemoteField = New RemoteField
    
    With NewRemoteField
        .Path = Path
        .Name = Name
        .Caption = Caption
        .CreationTime = Application.Evaluate("=NOW()")
        .ID = Left$(HashSHA1(.Path & "\" & .Name), HASH_LEN)
    End With
    
    Set CreateNewField = NewRemoteField
End Function
