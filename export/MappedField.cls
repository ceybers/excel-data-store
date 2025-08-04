VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MappedField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Version4.TableMap"
Option Explicit

Private Type TMappedField
    ListColumnName As String
    FieldID As String
    Selected As Boolean
End Type
Private This As TMappedField

Public Property Get ListColumnName() As String
    ListColumnName = This.ListColumnName
End Property

Public Property Let ListColumnName(ByVal RHS As String)
    This.ListColumnName = RHS
End Property

Public Property Get FieldID() As String
    FieldID = This.FieldID
End Property

Public Property Let FieldID(ByVal RHS As String)
    This.FieldID = RHS
End Property

Public Property Get Selected() As Boolean
    Selected = This.Selected
End Property

Public Property Let Selected(ByVal RHS As Boolean)
    This.Selected = RHS
End Property

Public Function Serialize() As String
    Dim Result As String
    Result = ListColumnName & Chr$(ASCII_US) & FieldID
    
    Serialize = Result
End Function

Public Sub Deserialize(ByVal SerializedString As String)
    Dim SplitString As Variant
    SplitString = Split(SerializedString, Chr$(ASCII_US))
    Debug.Assert UBound(SplitString) = 1
    
    ListColumnName = SplitString(0)
    FieldID = SplitString(1)
End Sub
