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
    Items As Collection
    KeyColumnName As String
    KeyPath As String
    KeyIDs As Collection
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

Public Sub AddMany(ByVal Values As Variant, ByVal RowOffset As Long)
    Dim i As Long
    For i = 1 To UBound(Values, 1)
        Debug.Assert VarType(Values(i, 1)) <> vbError
        Add Values(i, 1), i - RowOffset
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
    Log.Message "ResolveIDs()", "KeyMap"
    
    If Count = 0 Then Exit Sub
    
    Log.Message " KeyMap ResolveIDs", "KeyMap"
    Dim Keys As Variant
    ReDim Keys(1 To Count) As String
    
    CollectionClear This.KeyIDs
    
    Dim i As Long
    For i = 1 To Count
        Keys(i) = This.Items.Item(i).KeyValue
    Next i
    
    Dim KeyIDs As Variant
    Log.Message " Remote.Keys.GetIDsFromQuery", "KeyMap"
    KeyIDs = Remote.Keys.GetIDsFromQuery(KeyPath, Keys)
    
    Log.Message " For i; This.KeyIDs.Add", "KeyMap"
    For i = 1 To Count
        This.Items.Item(i).KeyID = KeyIDs(i)
        If KeyIDs(i) <> vbNullString Then
            ' FIX TODO Assumes Key Column has no duplicates. Crashes if not!
            On Error Resume Next
            This.KeyIDs.Add Item:=This.Items.Item(i).RowIndex, Key:=KeyIDs(i)
            On Error GoTo 0
        End If
    Next i
    
    Log.Message "ResolveIDs() End", "KeyMap"
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
