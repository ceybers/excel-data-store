VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Remote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "RemoteDataStore.Models"
Option Explicit

Private Const WS_KEYS_NAME As String = "Keys"
Private Const WS_FIELDS_NAME As String = "Fields"
Private Const WS_VALUES_NAME  As String = "Values"
Private Const WS_VALUES_HISTORY_NAME  As String = "ValuesHistory"
Private Const WS_COMMITS_NAME  As String = "Commits"
Private Const WS_MAPS_NAME  As String = "Maps"
Private Const WS_TABLES_NAME  As String = "Tables"

Private Type TState
    Workbook As Workbook
    RemoteKeys As RemoteKeys2
    RemoteFields As RemoteFields
    RemoteValues As RemoteValues2
    RemoteCommits As RemoteCommits
    RemoteMaps As RemoteMaps
    RemoteTables As RemoteTables
End Type
Private This As TState

Public Property Get IsValid() As Boolean
    If Not IsWorkbookOpen(This.Workbook) Then Exit Property
    IsValid = Not This.Workbook Is Nothing
End Property

Public Property Get Keys() As RemoteKeys2
    Log.Message "Remote.Keys Property Get", "Remote", Verbose_Level
    Set Keys = This.RemoteKeys
End Property

Public Property Get Fields() As RemoteFields
    Log.Message "Remote.Fields Property Get", "Remote", Verbose_Level
    If This.RemoteFields Is Nothing Then
        Set This.RemoteFields = New RemoteFields
        This.RemoteFields.Load This.Workbook.Worksheets.Item(WS_FIELDS_NAME)
    End If
    Set Fields = This.RemoteFields
End Property

Public Property Get Values() As RemoteValues2
    Log.Message "Remote.Values Property Get", "Remote", Verbose_Level
    If This.RemoteValues Is Nothing Then
        Set This.RemoteValues = New RemoteValues2
        This.RemoteValues.Load This.Workbook.Worksheets.Item(WS_VALUES_NAME), This.Workbook.Worksheets.Item(WS_VALUES_HISTORY_NAME)
    End If
    Set Values = This.RemoteValues
End Property

Public Property Get Commits() As RemoteCommits
    Log.Message "Remote.Commits Property Get", "Remote", Verbose_Level
    Set Commits = This.RemoteCommits
End Property

Public Property Get Maps() As RemoteMaps
    Log.Message "Remote.Maps Property Get", "Remote", Verbose_Level
    Set Maps = This.RemoteMaps
End Property

Public Property Get Tables() As RemoteTables
    Log.Message "Remote.Tables Property Get", "Remote", Verbose_Level
    If This.RemoteTables Is Nothing Then
        Set This.RemoteTables = New RemoteTables
        This.RemoteTables.Load This.Workbook.Worksheets.Item(WS_TABLES_NAME)
    End If
    Set Tables = This.RemoteTables
End Property

Public Sub Load(ByVal Workbook As Workbook)
    Log.Message "Remote.Load()", "Remote", Info_Level
    Set This.Workbook = Workbook
    
    Dim DEBUGRepoState As String
    DEBUGRepoState = "Repo State = " & _
        This.Workbook.Worksheets.Item(WS_KEYS_NAME).UsedRange.Rows.Count & "k " & _
        This.Workbook.Worksheets.Item(WS_FIELDS_NAME).UsedRange.Rows.Count & "f " & _
        This.Workbook.Worksheets.Item(WS_VALUES_NAME).UsedRange.Rows.Count & "v"
    
    Log.Message DEBUGRepoState, "Remote", Debug_Level
    
    Reload
End Sub

Public Sub Reload()
    Log.Message "Remote.Reload()", "Remote", Debug_Level
    
    With This
        Set .RemoteKeys = New RemoteKeys2
        .RemoteKeys.Load This.Workbook.Worksheets.Item(WS_KEYS_NAME)
        
        If Not .RemoteFields Is Nothing Then
            Set .RemoteFields = New RemoteFields
            .RemoteFields.Load This.Workbook.Worksheets.Item(WS_FIELDS_NAME)
        End If
        
        If Not .RemoteValues Is Nothing Then
            Set .RemoteValues = New RemoteValues2
            .RemoteValues.Load This.Workbook.Worksheets.Item(WS_VALUES_NAME), This.Workbook.Worksheets.Item(WS_VALUES_HISTORY_NAME)
        End If
        
        Set .RemoteMaps = New RemoteMaps
        .RemoteMaps.Load This.Workbook.Worksheets.Item(WS_MAPS_NAME)
        
        Set .RemoteTables = New RemoteTables
        .RemoteTables.Load This.Workbook.Worksheets.Item(WS_TABLES_NAME)
        
        Set .RemoteCommits = New RemoteCommits
        .RemoteCommits.Load This.Workbook.Worksheets.Item(WS_COMMITS_NAME)
    End With
    
    Log.Message "Remote.Reload() End", "Remote", Debug_Level
End Sub

Public Sub Show()
    This.Workbook.Windows.Item(1).Visible = True
End Sub

Public Sub Rebuild()
    RebuildRemoteTable This.Workbook.Worksheets.Item(WS_KEYS_NAME), _
        Array("KeyID", "Path", "Key", "Caption", "CreationTime", "ModificationTime", "DeletionTime")
    MsgBox MSG_REMOTE_REBUILD_KEYS, vbInformation + vbOKOnly, APP_TITLE
    
    RebuildRemoteTable This.Workbook.Worksheets.Item(WS_FIELDS_NAME), _
        Array("FieldID", "Path", "Name", "Caption", "CreationTime", "ModificationTime", "DeletionTime")
    MsgBox MSG_REMOTE_REBUILD_FIELDS, vbInformation + vbOKOnly, APP_TITLE
End Sub

Public Sub SaveWorkbook()
    This.Workbook.Save
End Sub

Public Sub CloseWorkbook()
    This.Workbook.Close
    With This
        Set .Workbook = Nothing
        Set .RemoteKeys = Nothing
        Set .RemoteFields = Nothing
        Set .RemoteValues = Nothing
        Set .RemoteCommits = Nothing
        Set .RemoteMaps = Nothing
        Set .RemoteTables = Nothing
    End With
End Sub
