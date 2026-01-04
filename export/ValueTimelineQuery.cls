VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueTimelineQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Version4.Queries")
Option Explicit

Private Type TValueTimelineQuery
    MappedTable As MappedTable
    Remote As Remote
    Results As Variant
    
    KeyIDs As Variant
    KeyID As String
    FieldIDs As Variant
    FieldID As String
End Type
Private This As TValueTimelineQuery

Public Property Get Results() As Variant
    Results = This.Results
End Property

Public Property Get MappedTable() As MappedTable
    Set MappedTable = This.MappedTable
End Property

Public Property Set MappedTable(ByVal RHS As MappedTable)
    Set This.MappedTable = RHS
End Property

Public Property Get Remote() As Remote
    Set Remote = This.Remote
End Property

Public Property Set Remote(ByVal RHS As Remote)
    Set This.Remote = RHS
End Property

Public Property Get CanRun() As Boolean
    This.KeyIDs = This.MappedTable.TableMap.KeyMap.GetKeyIDs
    If IsEmpty(This.KeyIDs) Then Exit Property
    
    This.KeyID = GetFirstNonEmptyString(This.KeyIDs)
    If This.KeyID = vbNullString Then Exit Property
    
    This.FieldIDs = This.MappedTable.TableMap.FieldMap.GetFieldIDs
    If IsEmpty(This.FieldIDs) Then Exit Property
    
    This.FieldID = GetFirstNonEmptyString(This.FieldIDs)
    If This.FieldID = vbNullString Then Exit Property
    
    CanRun = True
End Property

Public Sub Run()
    If Not CanRun Then
        This.Results = Empty
    Else
        This.Results = This.Remote.Values.GetAllValues(This.KeyID, This.FieldID)
    End If
End Sub

Private Function GetFirstNonEmptyString(ByVal ThisArray As Variant) As Variant
    Dim i As Long
    For i = 1 To UBound(ThisArray)
        If ThisArray(i) <> vbNullString Then
            GetFirstNonEmptyString = ThisArray(i)
            Exit Function
        End If
    Next i
End Function
