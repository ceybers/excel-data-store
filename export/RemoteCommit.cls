VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteCommit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder("RemoteDataStore.Models.Singular")
Option Explicit

Private Type TRemoteCommit
    CommitID As String
    TableID As String
    MapID As String
    Caption As String
    Timestamp As Date
    Message As String
End Type

Private This As TRemoteCommit

Public Property Get CommitID() As String
    CommitID = This.CommitID
End Property

Public Property Let CommitID(ByVal RHS As String)
    This.CommitID = RHS
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

Public Property Get Caption() As String
    Caption = This.Caption
End Property

Public Property Let Caption(ByVal RHS As String)
    This.Caption = RHS
End Property

Public Property Get Timestamp() As Date
    Timestamp = This.Timestamp
End Property

Public Property Let Timestamp(ByVal RHS As Date)
    This.Timestamp = RHS
End Property

Public Property Get Message() As String
    Message = This.Message
End Property

Public Property Let Message(ByVal RHS As String)
    This.Message = RHS
End Property

'@Ignore ParameterCanBeByVal
Public Sub LoadFromArray(ByRef VariantArray As Variant, ByVal Row As Long)
    With This
        .CommitID = VariantArray(Row, 1)
        .TableID = VariantArray(Row, 2)
        .MapID = VariantArray(Row, 3)
        .Caption = VariantArray(Row, 4)
        .Timestamp = VariantArray(Row, 5)
        .Message = VariantArray(Row, 6)
    End With
End Sub

'@Ignore ParameterCanBeByVal
Public Sub SaveToArray(ByRef VariantArray As Variant, ByVal Row As Long)
    With This
        VariantArray(Row, 1) = .CommitID
        VariantArray(Row, 2) = .TableID
        VariantArray(Row, 3) = .MapID
        VariantArray(Row, 4) = .Caption
        VariantArray(Row, 5) = .Timestamp
        VariantArray(Row, 6) = .Message
    End With
End Sub
