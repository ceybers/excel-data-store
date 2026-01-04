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
    KeyValues As Variant ' 1 to n
    KeyIDs As Variant ' 1 to n
    KeyIDtoRow As Variant ' 1 to n, 1 to 2
    
    KeyColumnName As String
    KeyPath As String
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
    If IsEmpty(This.KeyValues) Then
        Count = 0
        Exit Property
    End If
    
    Count = UBound(This.KeyValues)
End Property

Public Sub AddMany(ByVal Values As Variant)
    ReDim This.KeyValues(1 To UBound(Values))
    
    Dim i As Long
    For i = 1 To UBound(Values, 1)
        This.KeyValues(i) = Values(i, 1)
    Next i
End Sub

Public Sub ResolveIDs(ByVal Remote As Remote)
    Log.Message "ResolveIDs()... starts", "KeyMap", Info_Level
    
    If Count = 0 Then Exit Sub
    
    Log.Message " Remote.Keys.GetIDsFromQuery", "KeyMap", Verbose_Level
    This.KeyIDs = Remote.Keys.GetIDsFromQuery(KeyPath, This.KeyValues)
    
    ReDim This.KeyIDtoRow(1 To Count, 1 To 2)
    
    Dim i As Long
    For i = 1 To Count
        This.KeyIDtoRow(i, 1) = This.KeyIDs(i)
        This.KeyIDtoRow(i, 2) = i
    Next i
    
    ArraySort.QuickSort2 This.KeyIDtoRow
    
    Log.Message "ResolveIDs()... end", "KeyMap", Info_Level
End Sub

Public Function GetRowFromKeyID(ByVal KeyID As String) As Long
    Dim ResultIndex As Long
    ResultIndex = ArraySearch.BinarySearch2(This.KeyIDtoRow, KeyID)
    If ResultIndex = -1 Then Exit Function
    
    GetRowFromKeyID = This.KeyIDtoRow(ResultIndex, 2)
End Function

Public Function GetKeyIDs() As Variant
    GetKeyIDs = This.KeyIDs
End Function

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
