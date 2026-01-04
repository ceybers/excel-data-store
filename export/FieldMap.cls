VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.TableMap"
Option Explicit

Private Type TFieldMap
    Items As Collection
End Type
Private This As TFieldMap

Private Sub Class_Initialize()
    Set This.Items = New Collection
End Sub

Public Property Get Count() As Long
    Count = This.Items.Count
End Property

Public Property Get Item(ByVal Index As Long) As MappedField
    Set Item = This.Items.Item(Index)
End Property

Public Sub Clear()
    CollectionClear This.Items
End Sub

Public Sub Add(ByVal ListColumnName As String, ByVal FieldID As String)
    Dim NewMappedField As MappedField
    Set NewMappedField = New MappedField
    
    With NewMappedField
        .ListColumnName = ListColumnName
        .FieldID = FieldID
    End With
    
    This.Items.Add Item:=NewMappedField, Key:=FieldID
End Sub

Public Sub SelectField(ByVal ListColumnName As String)
    Dim ThisMappedField As MappedField
    For Each ThisMappedField In This.Items
        If ThisMappedField.ListColumnName = ListColumnName Then
            ThisMappedField.Selected = True
            Exit Sub
        End If
    Next ThisMappedField
End Sub

Public Sub SelectAllFields()
    Dim ThisMappedField As MappedField
    For Each ThisMappedField In This.Items
        ThisMappedField.Selected = True
    Next ThisMappedField
End Sub

Public Sub UnselectAllFields()
    Dim ThisMappedField As MappedField
    For Each ThisMappedField In This.Items
        ThisMappedField.Selected = False
    Next ThisMappedField
End Sub

Public Function Serialize() As String
    If Count = 0 Then Exit Function
    
    Dim Result() As Variant
    ReDim Result(1 To Count)
    
    Dim i As Long
    For i = 1 To Count
        Result(i) = This.Items.Item(i).Serialize
    Next i
    
    Serialize = Join$(Result, Chr$(ASCII_RS))
End Function

Public Sub Deserialize(ByVal SerializedString As String)
    CollectionClear This.Items
    
    Dim SplitString As Variant
    SplitString = Split(SerializedString, Chr$(ASCII_RS))
    
    Dim i As Long
    For i = 0 To UBound(SplitString)
        Dim MappedField As MappedField
        Set MappedField = New MappedField
        
        MappedField.Deserialize SplitString(i)
        
        This.Items.Add Item:=MappedField, Key:=MappedField.FieldID
    Next i
End Sub

Public Function GetFieldIDs() As Variant
    If Count = 0 Then Exit Function
    
    Dim Result As Variant
    ReDim Result(1 To Count) As String
    
    Dim Cursor As Long
    
    Dim ThisMappedField As MappedField
    For Each ThisMappedField In This.Items
        If ThisMappedField.Selected = True Then
            Cursor = Cursor + 1
            Result(Cursor) = ThisMappedField.FieldID
        End If
    Next ThisMappedField
    
    If Cursor = 0 Then Exit Function
    
    ReDim Preserve Result(1 To Cursor)
    
    GetFieldIDs = Result
End Function

Public Function GetByFieldID(ByVal FieldID As String) As MappedField
    Set GetByFieldID = This.Items.Item(FieldID)
End Function

'Public Function TryGetMappedFieldByName(ByVal Name As String, ByRef OutMappedField As MappedField) As Boolean
'    Dim ThisMappedField As MappedField
'    For Each ThisMappedField In This.Items
'        If ThisMappedField.ListColumnName = Name Then
'            Set OutMappedField = ThisMappedField
'            TryGetMappedFieldByName = True
'            Exit Function
'        End If
'    Next ThisMappedField
'End Function
