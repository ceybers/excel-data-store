VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MappedKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder("Version4.TableMap")
Option Explicit

Private Type TMappedKey
    KeyValue As String
    KeyID As String
    RowIndex As Long
End Type
Private This As TMappedKey

Public Property Get KeyValue() As String
    KeyValue = This.KeyValue
End Property

Public Property Let KeyValue(ByVal RHS As String)
    This.KeyValue = RHS
End Property

Public Property Get KeyID() As String
    KeyID = This.KeyID
End Property

Public Property Let KeyID(ByVal RHS As String)
    This.KeyID = RHS
End Property

Public Property Get RowIndex() As Long
    RowIndex = This.RowIndex
End Property

Public Property Let RowIndex(ByVal RHS As Long)
    This.RowIndex = RHS
End Property
