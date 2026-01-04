VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.TableMap"
Option Explicit

Private Type TTableMap
    Caption As String
    TableID As String
    MapID As String
    
    KeyMap As KeyMap
    FieldMap As FieldMap
End Type
Private This As TTableMap

Public Property Get IsMapped() As Boolean
    If This.KeyMap.KeyColumnName = vbNullString Then Exit Property
    If This.KeyMap.KeyPath = vbNullString Then Exit Property
    If This.FieldMap.Count = 0 Then Exit Property
    
    IsMapped = True
End Property

Public Property Get Caption() As String
    Caption = This.Caption
End Property

Public Property Get TableID() As String
    TableID = This.TableID
End Property

Public Property Let TableID(ByVal RHS As String)
    This.TableID = RHS
End Property

Public Property Get MapID() As String
    MapID = This.MapID
End Property

Public Property Let MapID(ByVal RHS As String)
    This.MapID = RHS
End Property

Public Property Get KeyMap() As KeyMap
    Set KeyMap = This.KeyMap
End Property

Public Property Get FieldMap() As FieldMap
    Set FieldMap = This.FieldMap
End Property

Private Sub Class_Initialize()
    Set This.KeyMap = New KeyMap
    Set This.FieldMap = New FieldMap
End Sub

Public Sub SetCaption(ByVal ListObjectName As String)
    This.Caption = This.KeyMap.KeyColumnName & " @ " & ListObjectName & " [" & This.FieldMap.Count & "]"
End Sub

Public Function Serialize() As String
    Serialize = This.KeyMap.Serialize & Chr$(ASCII_FS) & This.FieldMap.Serialize
End Function

Public Sub Deserialize(ByVal SerializedString As String)
    Dim SplitString As Variant
    SplitString = Split(SerializedString, Chr$(ASCII_FS))
    Debug.Assert UBound(SplitString) = 1
    
    This.KeyMap.Deserialize SplitString(0)
    This.FieldMap.Deserialize SplitString(1)
End Sub
