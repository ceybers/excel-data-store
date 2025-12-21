VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.TableMap"
Option Explicit

Private Type TKeyMap
    Items As Collection ' ? Collection<MappedKey>
    KeyColumnName As String
    KeyPath As String
    KeyIDs As Collection ' ? Collection<String>
End Type
Private This As TKeyMap

Public Property Get KeyColumnName() As String
    KeyColumnName = This.KeyColumnName
End Property

Public Property Let KeyColumnName(ByVal RHS As String)
    This.KeyColumnName = RHS
End Property

Public Property Get KeyPath() As String
    KeyPath = This.KeyPath
End Property

Public Property Let KeyPath(ByVal RHS As String)
    This.KeyPath = RHS
End Property

Public Property Get Count() As Long
    Count = This.Items.Count
End Property

Private Sub Class_Initialize()
    Set This.Items = New Collection
    Set This.KeyIDs = New Collection
End Sub

Public Sub Add(ByVal KeyValue As String, ByVal RowIndex As Long)
    If KeyValue = vbNullString Then Exit Sub
    
    Dim MappedKey As MappedKey
    Set MappedKey = New MappedKey
    With MappedKey
        .KeyValue = KeyValue
        .RowIndex = RowIndex
    End With
    This.Items.Add Item:=MappedKey
End Sub

Public Sub AddMany(ByVal Values As Variant)
    Dim i As Long
    For i = 1 To UBound(Values, 1)
        Debug.Assert VarType(Values(i, 1)) <> vbError
        Add Values(i, 1), i
    Next i
End Sub

Public Function Serialize() As String
    Dim Result As String
    Result = KeyColumnName & Chr$(ASCII_US) & KeyPath
    
    Serialize = Result
End Function

Public Sub Deserialize(ByVal SerializedString As String)
    Dim SplitString As Variant
    SplitString = Split(SerializedString, Chr$(ASCII_US))
    Debug.Assert UBound(SplitString) = 1
    
    KeyColumnName = SplitString(0)
    KeyPath = SplitString(1)
End Sub

Public Sub ResolveIDs(ByVal Remote As Remote)
    Log.Message "ResolveIDs()... starts", "KeyMap", Info_Level
    
    If Count = 0 Then Exit Sub
    
    Log.Message " KeyMap ResolveIDs", "KeyMap"
    Dim Keys As Variant
    ReDim Keys(1 To Count) As String
    
    CollectionClear This.KeyIDs
    Log.Message " This.KeyIDs cleared", "KeyMap", Verbose_Level
    
    Log.Message " For i = 1 to " & Count, "KeyMap", Verbose_Level
    Log.Message "  Keys(i) = This.Items.Item(i).KeyValue", "KeyMap", Verbose_Level
    Dim i As Long
    
    
    Log.Message "Test Keys(i)... starts", "KeyMap", Error_Level
    For i = 1 To Count
        Keys(i) = Keys(i)
    Next i
    Log.Message "Test Keys(i)... done", "KeyMap", Error_Level
    
    Log.Message "Test This.Items.Item(i).KeyValue... starts", "KeyMap", Error_Level
    For i = 1 To Count
        This.Items.Item(i).KeyValue = This.Items.Item(i).KeyValue
    Next i
    ' TODO FIX This loop takes 200ms for 1000 items!!!
    Log.Message "Test This.Items.Item(i).KeyValue... done", "KeyMap", Error_Level
    
    For i = 1 To Count
        ' Keys(i) = Variant/String(1 to 1002)
        ' This.Items = Collection<MappedKey>
        Keys(i) = This.Items.Item(i).KeyValue
    Next i
    Log.Message " Next i", "KeyMap", Verbose_Level
    
    Dim KeyIDs As Variant
    Log.Message " Remote.Keys.GetIDsFromQuery", "KeyMap", Verbose_Level
    KeyIDs = Remote.Keys.GetIDsFromQuery(KeyPath, Keys)
    
    Log.Message " For i = 1 to " & Count & "; This.KeyIDs.Add", "KeyMap", Verbose_Level
    For i = 1 To Count
        This.Items.Item(i).KeyID = KeyIDs(i)
        If KeyIDs(i) <> vbNullString Then
            ' FIX TODO Assumes Key Column has no duplicates. Crashes if not!
            'On Error Resume Next
            This.KeyIDs.Add Item:=This.Items.Item(i).RowIndex, Key:=KeyIDs(i)
            'On Error GoTo 0
        End If
        
        'If i Mod 100 = 0 Then Log.Message "  " & i, "KeyMap", Verbose_Level
    Next i
    
    Log.Message "ResolveIDs()... end", "KeyMap", Info_Level
End Sub

Public Function GetRowFromKeyID(ByVal KeyID As String) As Long
    GetRowFromKeyID = This.KeyIDs.Item(KeyID)
End Function

Public Function GetKeyIDs() As Variant
    If Count = 0 Then Exit Function
    
    Dim Result As Variant
    ReDim Result(1 To Count) As String
    
    Dim i As Long
    For i = 1 To Count
        Result(i) = This.Items.Item(i).KeyID
    Next i
    
    GetKeyIDs = Result
End Function
