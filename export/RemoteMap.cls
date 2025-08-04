VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "RemoteDataStore.Models.Singular"
Option Explicit

Private Type TRemoteMap
    ID As String
    SerializedString As String
    Caption As String
    CreationTime As Date
    ModificationTime As Date
    DeletionTime As Date
End Type
Private This As TRemoteMap

Public Property Get ID() As String
    ID = This.ID
End Property

Public Property Let ID(ByVal RHS As String)
    This.ID = RHS
End Property

Public Property Get SerializedString() As String
    SerializedString = This.SerializedString
End Property

Public Property Let SerializedString(ByVal RHS As String)
    This.SerializedString = RHS
End Property

Public Property Get Caption() As String
    Caption = This.Caption
End Property

Public Property Let Caption(ByVal RHS As String)
    This.Caption = RHS
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
        .SerializedString = VariantArray(Row, 2)
        .Caption = VariantArray(Row, 3)
        .CreationTime = VariantArray(Row, 4)
        .ModificationTime = VariantArray(Row, 5)
        .DeletionTime = VariantArray(Row, 6)
    End With
End Sub

'@Ignore ParameterCanBeByVal
Public Sub SaveToArray(ByRef VariantArray As Variant, ByVal Row As Long)
    With This
        VariantArray(Row, 1) = .ID
        VariantArray(Row, 2) = .SerializedString
        VariantArray(Row, 3) = .Caption
        VariantArray(Row, 4) = .CreationTime
        VariantArray(Row, 5) = .ModificationTime
        VariantArray(Row, 6) = .DeletionTime
    End With
End Sub
