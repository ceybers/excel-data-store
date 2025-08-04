VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "RemoteDataStore.Models.Singular"
Option Explicit

Private Type TRemoteValue
    ID As String
    KeyID As String
    FieldID As String
    Value As Variant
    Latest As String
    CreationTime As Date
    ModificationTime As Date
    DeletionTime As Date
    CommitID As String
End Type
Private This As TRemoteValue

Public Property Get ID() As String
    ID = This.ID
End Property

Public Property Let ID(ByVal RHS As String)
    This.ID = RHS
End Property

Public Property Get KeyID() As String
    KeyID = This.KeyID
End Property

Public Property Let KeyID(ByVal RHS As String)
    This.KeyID = RHS
End Property

Public Property Get FieldID() As String
    FieldID = This.FieldID
End Property

Public Property Let FieldID(ByVal RHS As String)
    This.FieldID = RHS
End Property

Public Property Get Value() As Variant
    Value = This.Value
End Property

Public Property Let Value(ByVal RHS As Variant)
    This.Value = RHS
End Property

Public Property Get Latest() As String
    Latest = This.Latest
End Property

Public Property Let Latest(ByVal RHS As String)
    This.Latest = RHS
End Property

Public Property Get CreationTime() As Date
    CreationTime = This.CreationTime
End Property

Public Property Let CreationTime(ByVal RHS As Date)
    This.CreationTime = RHS
End Property

Public Property Get ModificationTime() As Date
    ModificationTime = This.ModificationTime
End Property

Public Property Let ModificationTime(ByVal RHS As Date)
    This.ModificationTime = RHS
End Property

Public Property Get DeletionTime() As Date
    DeletionTime = This.DeletionTime
End Property

Public Property Let DeletionTime(ByVal RHS As Date)
    This.DeletionTime = RHS
End Property

Public Property Get CommitID() As String
    CommitID = This.CommitID
End Property

Public Property Let CommitID(ByVal RHS As String)
    This.CommitID = RHS
End Property

'@Ignore ParameterCanBeByVal
Public Sub LoadFromArray(ByRef VariantArray As Variant, ByVal Row As Long)
    With This
        .ID = VariantArray(Row, 1)
        .KeyID = VariantArray(Row, 2)
        .FieldID = VariantArray(Row, 3)
        .Value = VariantArray(Row, 4)
        .Latest = VariantArray(Row, 5)
        .CreationTime = VariantArray(Row, 6)
        .ModificationTime = VariantArray(Row, 7)
        .DeletionTime = VariantArray(Row, 8)
        .CommitID = VariantArray(Row, 9)
    End With
End Sub

'@Ignore ParameterCanBeByVal
Public Sub SaveToArray(ByRef VariantArray As Variant, ByVal Row As Long)
    With This
        VariantArray(Row, 1) = .ID
        VariantArray(Row, 2) = .KeyID
        VariantArray(Row, 3) = .FieldID
        VariantArray(Row, 4) = .Value
        VariantArray(Row, 5) = .Latest
        VariantArray(Row, 6) = .CreationTime
        VariantArray(Row, 7) = .ModificationTime
        VariantArray(Row, 8) = .DeletionTime
        VariantArray(Row, 9) = .CommitID
    End With
End Sub
