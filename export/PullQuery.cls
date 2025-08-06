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

Private Const DO_DEBUG As Boolean = True

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
    
    Dim FieldIDs As Variant
    Log.Message "PullQuery.Run GetFields", "PullQuery", Info_Level
    FieldIDs = This.MappedTable.GetFieldIDs
    If IsEmpty(FieldIDs) Then Exit Sub
    
    Dim ExistingValues As Variant
    ExistingValues = This.MappedTable.GetExistingValuesFromTable(KeyIDs, FieldIDs)
    
    Dim ChangedValues As Variant
    ChangedValues = This.Remote.Values.GetChangeMask(KeyIDs, FieldIDs, ExistingValues)
    
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangedValues, hcBeforePull
    
    Dim ChangeCount As Long
    ChangeCount = This.Remote.Values.GetChangeCount(KeyIDs, FieldIDs, ExistingValues)
    
    If ChangeCount = 0 Then
        MsgBox MSG_PULL_NOCHANGES, vbInformation + vbOKOnly, APP_TITLE
    ElseIf vbNo = MsgBox(Replace$(MSG_PULL_CONFIRM, "{0}", ChangeCount), vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) Then
        Exit Sub
    End If
    
    RangeHighlighter.RemoveHighlights This.MappedTable.ListObject
    This.MappedTable.HighlightChanges KeyIDs, FieldIDs, ChangedValues, hcAfterPull
    
    Dim NewValues As Variant
    Log.Message "PullQuery.Run GetValues", "PullQuery", Info_Level
    NewValues = This.Remote.Values.GetValues(KeyIDs, FieldIDs)
    Log.Message "PullQuery.Run UpdateFromRemote", "PullQuery", Info_Level
    UpdateFromRemote KeyIDs, FieldIDs, NewValues
    
    Log.Message "PullQuery.Run() End", "PullQuery", Info_Level
End Sub

Private Sub UpdateFromRemote(ByVal KeyIDs As Variant, ByVal FieldIDs As Variant, ByVal NewValues As Variant)
    Log.Message "PullQuery.UpdateFromRemote creating KeyIndexes", "PullQuery", Info_Level
    Dim KeyIndexes As Variant
    ReDim KeyIndexes(1 To UBound(KeyIDs))
    Dim i As Long
    For i = 1 To UBound(KeyIDs)
        KeyIndexes(i) = This.MappedTable.GetRowIndexFromKeyID(KeyIDs(i))
    Next i
    
    Log.Message "PullQuery.UpdateFromRemote Looping through Fields", "PullQuery", Info_Level
    Dim f As Long
    For f = 1 To UBound(FieldIDs)
        Dim ListColumn As ListColumn
        Set ListColumn = This.MappedTable.GetListColumnFromFieldID(FieldIDs(f))
        
        Dim vv As Variant
        vv = ListColumn.DataBodyRange.Value2
        
        Dim k As Long
        For k = 1 To UBound(KeyIndexes)
            vv(KeyIndexes(k), 1) = NewValues(k, f)
        Next k
        
        Log.Message "PullQuery.UpdateFromRemote writing to " & ListColumn.DataBodyRange.Address(False, False), "PullQuery", Info_Level
        ListColumn.DataBodyRange.Value2 = vv
    Next f
End Sub
