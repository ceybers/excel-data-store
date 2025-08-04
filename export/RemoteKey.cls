VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "RemoteDataStore.Models.Singular"
Option Explicit

Private Type TRemoteKey
    ID As String
    Path As String
    Key As String
    Synonym As String
    CreationTime As Date
    ModificationTime As Date
    DeletionTime As Date
End Type
Private This As TRemoteKey

Public Property Get ID() As String
    ID = This.ID
End Property

Public Property Let ID(ByVal RHS As String)
    This.ID = RHS
End Property

Public Property Get Path() As String
    Path = This.Path
End Property

Public Property Let Path(ByVal RHS As String)
    This.Path = RHS
End Property

Public Property Get Key() As String
    Key = This.Key
End Property

Public Property Let Key(ByVal RHS As String)
    This.Key = RHS
End Property

Public Property Get Synonym() As String
    Synonym = This.Synonym
End Property

Public Property Let Synonym(ByVal RHS As String)
    This.Synonym = RHS
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

'@Ignore ParameterCanBeByVal
Public Sub LoadFromArray(ByRef VariantArray As Variant, ByVal Row As Long)
    With This
        .ID = VariantArray(Row, 1)
        .Path = VariantArray(Row, 2)
        .Key = VariantArray(Row, 3)
        .Synonym = VariantArray(Row, 4)
        .CreationTime = VariantArray(Row, 5)
        .ModificationTime = VariantArray(Row, 6)
        .DeletionTime = VariantArray(Row, 7)
    End With
End Sub

'@Ignore ParameterCanBeByVal
Public Sub SaveToArray(ByRef VariantArray As Variant, ByVal Row As Long)
    With This
        VariantArray(Row, 1) = .ID
        VariantArray(Row, 2) = .Path
        VariantArray(Row, 3) = .Key
        VariantArray(Row, 4) = .Synonym
        VariantArray(Row, 5) = .CreationTime
        VariantArray(Row, 6) = .ModificationTime
        VariantArray(Row, 7) = .DeletionTime
    End With
End Sub
