VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PullQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "Version4.Queries"
Option Explicit

Private Type TPullQuery
    MappedTable As MappedTable
    Remote As Remote
End Type
Private This As TPullQuery

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

Public Sub Run()
    Log.Message "PullQuery.Run()", "PullQuery", Info_Level
    
    Log.Message "This.Remote.Load", "PullQuery", Info_Level
    This.Remote.Reload
    
    RangeHighlighter.RemoveHighlights This.MappedTable.ListObject
    
    Dim KeyIDs As Variant
    Log.Message "This.MappedTable.GetKeyIDs", "PullQuery", Info_Level
    KeyIDs = This.MappedTable.GetKeyIDs
    If IsEmpty(KeyIDs) Then Exit Sub
    
    Dim FieldIDs As Variant
    Log.Message "Run GetFields", "PullQuery", Info_Level
    FieldIDs = This.MappedTable.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    
    Dim LocalValues As Variant
    Log.Message "Run MappedTable.GetExistingValuesFromTable", "PullQuery", Info_Level
    LocalValues = This.MappedTable.GetExistingValuesFromTable(KeyIDs, FieldIDs)
    
    Dim RemoteValues As Variant
    Log.Message "PullQuery.Run GetValues", "PullQuery", Info_Level
    RemoteValues = This.Remote.Values.GetValues(KeyIDs, FieldIDs)
    
    Dim ChangeMask As Variant
    ChangeMask = GetChangeMask(LocalValues, RemoteValues)
    
    Dim ChangeCount As Long
    ChangeCount = CountChanges(ChangeMask)
    
    If ChangeCount = 0 Then
        MsgBox MSG_PULL_NOCHANGES, vbInformation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangeMask, hcBeforePull
    
    If vbNo = MsgBox(Replace$(MSG_PULL_CONFIRM, "{0}", ChangeCount), vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) Then
        Exit Sub
    End If
    
    RangeHighlighter.RemoveHighlights This.MappedTable.ListObject
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangeMask, hcAfterPull
    
    Log.Message "PullQuery.Run UpdateFromRemote", "PullQuery", Info_Level
    UpdateFromRemote KeyIDs, FieldIDs, RemoteValues
    
    Log.Message "PullQuery.Run() End", "PullQuery", Info_Level
End Sub

Private Sub UpdateFromRemote(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, ByVal NewValues As Variant)
    Log.Message "PullQuery.UpdateFromRemote creating KeyIndexes", "PullQuery", Info_Level
    
    Dim KeyIndexes As Variant
    ReDim KeyIndexes(1 To UBound(KeyIDs)) As Long
    
    Dim i As Long
    For i = 1 To UBound(KeyIDs)
        If KeyIDs(i) <> vbNullString Then
            KeyIndexes(i) = This.MappedTable.GetRowIndexFromKeyID(KeyIDs(i))
        End If
    Next i
    
    Log.Message "PullQuery.UpdateFromRemote Looping through Fields", "PullQuery", Info_Level
    Dim f As Long
    For f = 1 To UBound(FieldIDs)
        Dim ListColumn As ListColumn
        Set ListColumn = This.MappedTable.GetListColumnFromFieldID(FieldIDs(f))
        
        Dim vv As Variant
        vv = ListColumn.DataBodyRange.Value2
        If Not IsArray(vv) Then
            vv = VariantToArray(vv)
        End If
        
        Dim k As Long
        For k = 1 To UBound(KeyIndexes)
            If IsEmpty(KeyIndexes(k)) = False Then
                If KeyIndexes(k) > 0 Then
                    vv(KeyIndexes(k), 1) = NewValues(k, f)
                End If
            End If
        Next k
        
        RangeSetValueFromVariant ListColumn.DataBodyRange, vv
    Next f
End Sub
